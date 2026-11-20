"""Image memory regression tests. Requires Playwright, Pillow, and Microsoft Edge."""

import base64
import io
import unittest
from pathlib import Path
from urllib.parse import urlsplit

from PIL import Image, ImageDraw
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent.parent
BASE = "https://jta.test/"


def photo(width, height, orientation=1):
    image = Image.new("RGB", (width, height), "red")
    draw = ImageDraw.Draw(image)
    draw.rectangle((width // 2, 0, width, height // 2), fill="green")
    draw.rectangle((0, height // 2, width // 2, height), fill="blue")
    draw.rectangle((width // 2, height // 2, width, height), fill="yellow")
    output = io.BytesIO()
    exif = Image.Exif()
    exif[274] = orientation
    image.save(output, format="JPEG", quality=95, exif=exif)
    image.close()
    return output.getvalue()


class ImageMemoryTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.playwright = sync_playwright().start()
        cls.browser = cls.playwright.chromium.launch(channel="msedge", headless=True)
        cls.large_photo = photo(6000, 4000)

    @classmethod
    def tearDownClass(cls):
        cls.browser.close()
        cls.playwright.stop()

    def setUp(self):
        self.context = self.browser.new_context(service_workers="block")
        self.context.route(BASE + "**", self.serve)
        self.page = self.context.new_page()
        self.errors = []
        self.page.on("pageerror", lambda error: self.errors.append(str(error)))
        self.page.goto(BASE, wait_until="networkidle")
        self.page.wait_for_function("authState === 'signed_out'")

    def tearDown(self):
        self.context.close()
        self.assertEqual(self.errors, [])

    def serve(self, route):
        relative = urlsplit(route.request.url).path.lstrip("/") or "index.html"
        target = ROOT / relative
        self.assertTrue(target.resolve().is_relative_to(ROOT.resolve()))
        if relative == "js/vendor/msal-browser.min.js":
            route.fulfill(content_type="application/javascript", body="""
                window.msal = { PublicClientApplication: class {
                  async initialize() {}
                  async handleRedirectPromise() { return null; }
                  getAllAccounts() { return []; }
                }};
            """)
        elif target.is_file():
            route.fulfill(path=str(target))
        else:
            route.abort()

    def choose(self, data=None):
        previous = self.page.evaluate("imageLoadId")
        self.page.set_input_files("#file-input", {
            "name": "photo.jpg", "mimeType": "image/jpeg",
            "buffer": data if data is not None else self.large_photo,
        })
        self.page.wait_for_function(
            "previous => imageLoadId > previous && imgLoaded",
            arg=previous,
        )

    def set_limit(self, limit):
        self.page.click("#btn-settings")
        self.page.select_option("#cfg-image-max-edge", str(limit))
        self.page.click("#btn-save-settings")
        self.page.wait_for_function("authState === 'signed_out'")

    def crop(self, x=100, y=100, width=300, height=200):
        data = self.page.evaluate("""([x, y, w, h]) => {
            selection = {x, y, w, h};
            drawSelectionOverlay();
            return getCroppedDataUrl();
        }""", [x, y, width, height])
        return Image.open(io.BytesIO(base64.b64decode(data.split(",")[1])))

    def test_preview_memory_and_original_resolution_rotation(self):
        self.choose()
        dimensions = self.page.evaluate("""({
            source: [imageWidth, imageHeight],
            photo: [els.photoCanvas.width, els.photoCanvas.height],
            overlay: [els.selectionCanvas.width, els.selectionCanvas.height]
        })""")
        self.assertEqual(dimensions["source"], [6000, 4000])
        self.assertLessEqual(max(dimensions["photo"]), 1600)
        self.assertEqual(dimensions["photo"], dimensions["overlay"])
        # Before this fix the two canvases alone consumed 192 MB for this image.
        preview_bytes = dimensions["photo"][0] * dimensions["photo"][1] * 4 * 2
        self.assertLess(preview_bytes, 192_000_000 / 10)
        for expected in [(255, 0, 0), (0, 0, 255), (255, 255, 0), (0, 128, 0)]:
            with self.crop(width=1200, height=600) as crop:
                self.assertEqual(crop.size, (1200, 600))
                pixel = crop.getpixel((600, 300))
                for actual, wanted in zip(pixel, expected):
                    self.assertLess(abs(actual - wanted), 10)
            self.page.click("#btn-rotate-right")
        self.assertEqual(self.page.evaluate("rotationDegrees"), 0)
        # Exercise the real analysis request builder without sending test photos to Azure.
        encoded = self.page.evaluate("""async () => {
            selection = {x: 100, y: 100, w: 1200, h: 600};
            getAccessToken = async () => 'test-token';
            let sentImage;
            window.fetch = async (url, request) => {
                sentImage = JSON.parse(request.body).messages[0].content[1].image_url.url;
                return {ok: true, json: async () => ({choices: [{
                    finish_reason: 'stop', message: {content: JSON.stringify({
                        contentType: 'reading', sourceText: 'test', translation: 'test',
                        words: [], kanjiBreakdown: [], grammar: []
                    })}
                }]})};
            };
            await analyze();
            return sentImage;
        }""")
        with Image.open(io.BytesIO(base64.b64decode(encoded.split(",")[1]))) as sent:
            self.assertEqual(sent.size, (1200, 600))

    def test_pointer_mapping_and_resize_keep_crop(self):
        self.choose()
        self.page.locator("#selection-canvas").scroll_into_view_if_needed()
        box = self.page.locator("#selection-canvas").bounding_box()
        self.page.mouse.move(box["x"] + box["width"] * .1, box["y"] + box["height"] * .1)
        self.page.mouse.down()
        self.page.mouse.move(box["x"] + box["width"] * .4, box["y"] + box["height"] * .4, steps=5)
        self.page.mouse.up()
        selected = self.page.evaluate("selection")
        self.assertAlmostEqual(selected["x"], 600, delta=10)
        self.assertAlmostEqual(selected["y"], 400, delta=10)
        self.assertAlmostEqual(selected["w"], 1800, delta=10)
        self.assertAlmostEqual(selected["h"], 1200, delta=10)
        self.assertTrue(self.page.is_enabled("#btn-analyze"))
        self.page.set_viewport_size({"width": 390, "height": 800})
        self.page.wait_for_function("els.photoCanvas.style.width === '358px'")
        self.assertEqual(self.page.evaluate("selection"), selected)
        self.assertTrue(self.page.is_enabled("#btn-analyze"))
        self.page.evaluate("""() => {
            const rect = els.selectionCanvas.getBoundingClientRect();
            startDrag({clientX: rect.x + 20, clientY: rect.y + 20, preventDefault() {}});
            moveDrag({clientX: rect.right + 500, clientY: rect.bottom + 500, preventDefault() {}});
            endDrag();
        }""")
        with self.crop(
            **dict(zip(["x", "y", "width", "height"],
                       self.page.evaluate("[selection.x, selection.y, selection.w, selection.h]")))
        ) as crop:
            self.assertGreater(crop.width, 0)
            self.assertLessEqual(crop.width, 6000)

    def test_resolution_setting_is_opt_in_and_applies_to_next_photo(self):
        self.choose()
        self.set_limit(2048)
        self.assertEqual(self.page.evaluate("[imageWidth, imageHeight]"), [6000, 4000])
        self.choose()
        self.assertEqual(self.page.evaluate("[imageWidth, imageHeight]"), [2048, 1365])
        self.assertTrue(self.page.evaluate("typeof img.close === 'function'"))
        with self.crop() as crop:
            self.assertEqual(crop.size, (300, 200))
        self.page.evaluate("window.oldBitmap = img")
        self.choose(photo(800, 600))
        self.assertEqual(self.page.evaluate("[oldBitmap.width, oldBitmap.height]"), [0, 0])
        self.assertEqual(self.page.evaluate("[imageWidth, imageHeight]"), [800, 600])
        self.page.reload(wait_until="networkidle")
        self.page.click("#btn-settings")
        self.assertEqual(self.page.input_value("#cfg-image-max-edge"), "2048")
        self.page.select_option("#cfg-language", "ja")
        self.assertEqual(self.page.locator("[data-i18n='image.title']").inner_text(), "画像のメモリ使用量")
        self.assertNotIn("image.note", self.page.locator("[data-i18n='image.note']").inner_text())

    def test_small_photos_fallback_and_authentication_crop_restore(self):
        self.set_limit(4096)
        self.page.evaluate("window.createImageBitmap = undefined")
        self.choose()
        self.assertEqual(self.page.evaluate("[imageWidth, imageHeight]"), [4096, 2731])
        self.assertEqual(self.page.evaluate("img.tagName"), "CANVAS")
        self.page.evaluate("""async () => {
            selection = {x: 100, y: 100, w: 1200, h: 600};
            els.levelSelect.value = 'N4';
            els.analysisModeSelect.value = 'exercise';
            await persistPendingAnalysis();
            localStorage.setItem('jta_analysis_auth_pending', '1');
        }""")
        self.page.reload(wait_until="networkidle")
        self.page.wait_for_function("imgLoaded")
        self.assertEqual(self.page.evaluate("[imageWidth, imageHeight]"), [1200, 600])
        self.assertEqual(self.page.evaluate("selection"), {"x": 0, "y": 0, "w": 1200, "h": 600})
        self.assertEqual(self.page.input_value("#level-select"), "N4")
        self.assertEqual(self.page.input_value("#analysis-mode-select"), "exercise")
        self.assertTrue(self.page.is_enabled("#btn-analyze"))

    def test_camera_exif_orientation_with_and_without_resolution_limit(self):
        portrait = photo(6000, 4000, orientation=6)
        self.choose(portrait)
        self.assertEqual(self.page.evaluate("[imageWidth, imageHeight]"), [4000, 6000])
        with self.crop() as crop:
            self.assertGreater(crop.getpixel((150, 100))[2], 240)
        self.set_limit(2048)
        self.choose(portrait)
        self.assertEqual(self.page.evaluate("[imageWidth, imageHeight]"), [1365, 2048])
        with self.crop() as crop:
            self.assertGreater(crop.getpixel((150, 100))[2], 240)

    def test_failed_and_superseded_loads_release_sources(self):
        self.choose()
        self.page.evaluate("window.oldImage = img")
        self.page.set_input_files("#file-input", {
            "name": "broken.jpg", "mimeType": "image/jpeg", "buffer": b"not an image",
        })
        self.page.wait_for_selector(".error-box")
        self.assertFalse(self.page.evaluate("imgLoaded"))
        self.assertEqual(self.page.evaluate("oldImage.getAttribute('src')"), "")
        self.assertEqual(self.page.evaluate("[els.photoCanvas.width, els.selectionCanvas.width]"), [1, 1])
        self.assertTrue(self.page.is_disabled("#btn-analyze"))
        self.choose(photo(800, 600))
        self.assertEqual(self.page.locator(".error-box").count(), 0)
        results = self.page.evaluate("""async () => {
            const first = loadImageSource('data:image/jpeg;base64,bad')
                .then(() => 'loaded', error => error.name);
            const second = loadImageDataUrl(els.photoCanvas.toDataURL());
            return [await first, await second, imgLoaded];
        }""")
        self.assertEqual(results[0], "AbortError")
        self.assertTrue(results[2])


if __name__ == "__main__":
    unittest.main()
