"""Tests for the pure scaling / orientation math."""

from pdf_folder_print.scaling import fit_page, render_zoom


class TestRenderZoom:
    def test_dpi_override_wins(self):
        assert render_zoom(595, 842, 4961, 7016, dpi_override=300) == 300 / 72

    def test_native_zoom_covers_longer_axis(self):
        # A4 portrait in points; printer 300 dpi portrait pixels
        zoom = render_zoom(595, 842, 2480, 3508)
        # covers whichever axis needs the larger factor
        assert zoom == max(2480 / 595, 3508 / 842)


class TestFitPage:
    def test_portrait_page_portrait_image_no_rotation(self):
        # 1000x500 bitmap on 1000x1000 page: fills width exactly
        out_w, out_h, x, y, rotated = fit_page(1000, 500, 1000, 1000)
        assert (out_w, out_h) == (1000, 500)
        assert (x, y) == (0, 250)
        assert rotated is False

    def test_landscape_image_rotates_when_better(self):
        # 2000x1000 image on 800x1200 page:
        #   normal:   min(800/2000, 1200/1000) = 0.4  ->  800x400
        #   rotated:  min(800/1000, 1200/2000) = 0.6  ->  600x1200  (better fill)
        out_w, out_h, x, y, rotated = fit_page(2000, 1000, 800, 1200)
        assert rotated is True
        # rotated bitmap is 1000x2000 -> scaled by 0.6 -> 600x1200
        assert (out_w, out_h) == (600, 1200)
        assert (x, y) == (100, 0)

    def test_equal_scales_prefer_no_rotation(self):
        # Square image on square page: both scales equal -> no rotation
        out_w, out_h, x, y, rotated = fit_page(500, 500, 1000, 1000)
        assert (out_w, out_h) == (1000, 1000)
        assert (x, y) == (0, 0)
        assert rotated is False

    def test_oversized_page_scales_down(self):
        # A3-like bitmap on A4-like page: must scale down, aspect preserved
        out_w, out_h, x, y, rotated = fit_page(1191, 1684, 4961, 7016)
        scale_w = out_w / 1191
        scale_h = out_h / 1684
        assert abs(scale_w - scale_h) / scale_w < 0.01  # int() truncation tolerance
        assert out_w <= 4961 and out_h <= 7016

    def test_undersized_page_scales_up(self):
        # A5-like bitmap on A4-like page: must scale up, aspect preserved
        out_w, out_h, x, y, rotated = fit_page(595, 842, 4961, 7016)
        scale_w = out_w / 595
        scale_h = out_h / 842
        assert abs(scale_w - scale_h) / scale_w < 0.01  # int() truncation tolerance
        assert out_w > 595 and out_h > 842

    def test_centered_on_page(self):
        out_w, out_h, x, y, rotated = fit_page(1000, 500, 1200, 900)
        assert x == (1200 - out_w) // 2
        assert y == (900 - out_h) // 2
