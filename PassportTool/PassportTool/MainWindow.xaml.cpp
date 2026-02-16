#include "pch.h"
#include "MainWindow.xaml.h"
#if __has_include("MainWindow.g.cpp")
#include "MainWindow.g.cpp"
#endif
#include <cmath>
#include <algorithm>
#include <vector>

using namespace winrt;
using namespace Microsoft::UI::Xaml;
using namespace Microsoft::UI::Xaml::Controls;
using namespace Microsoft::UI::Xaml::Media;
using namespace Microsoft::UI::Xaml::Media::Imaging;
using namespace Microsoft::UI::Xaml::Input;
using namespace winrt::Windows::Storage;
using namespace winrt::Windows::Storage::Pickers;
using namespace winrt::Windows::Graphics::Imaging;
using namespace winrt::Windows::Storage::Streams;
using namespace winrt::Windows::ApplicationModel::DataTransfer;
using namespace winrt::Windows::Globalization::NumberFormatting;

namespace winrt::PassportTool::implementation
{
    // ──────────────────────────────────────────────────────────────
    // Construction & Initialisation
    // ──────────────────────────────────────────────────────────────

    MainWindow::MainWindow()
    {
        InitializeComponent();
        m_isLoaded = false;

        // Number-box formatting: 2 decimal places, round half-up
        DecimalFormatter formatter;
        formatter.IntegerDigits(1);
        formatter.FractionDigits(2);
        IncrementNumberRounder rounder;
        rounder.Increment(0.01);
        rounder.RoundingAlgorithm(RoundingAlgorithm::RoundHalfUp);
        formatter.NumberRounder(rounder);

        if (NbGap())    NbGap().NumberFormatter(formatter);
        if (NbSheetW()) NbSheetW().NumberFormatter(formatter);
        if (NbSheetH()) NbSheetH().NumberFormatter(formatter);
        if (NbImageW()) NbImageW().NumberFormatter(formatter);
        if (NbImageH()) NbImageH().NumberFormatter(formatter);

        // Slider -> ScrollViewer zoom
        if (ZoomSlider())
        {
            ZoomSlider().ValueChanged([this](auto&&, auto&& args)
            {
                if (m_zoomingFromMouse) return;
                m_zoomingFromSlider = true;
                if (CropScrollViewer())
                    CropScrollViewer().ChangeView(nullptr, nullptr, static_cast<float>(args.NewValue()), true);
                m_zoomingFromSlider = false;
            });
        }

        // ScrollViewer -> Slider sync
        if (CropScrollViewer())
        {
            CropScrollViewer().ViewChanged([this](auto&&, auto&&)
            {
                if (m_zoomingFromSlider || m_zoomingFromMouse) return;
                if (CropScrollViewer() && ZoomSlider())
                    ZoomSlider().Value(CropScrollViewer().ZoomFactor());
            });
        }
    }

    void MainWindow::Log([[maybe_unused]] winrt::hstring const& message)
    {
#ifdef _DEBUG
        std::wstring msg = L"[PassportTool] " + std::wstring(message) + L"\r\n";
        OutputDebugStringW(msg.c_str());
#endif
    }

    void MainWindow::OnWindowLoaded(IInspectable const&, RoutedEventArgs const&)
    {
        m_isLoaded = true;
        UpdateSheetSize();
        RefitCropContainer();
        UpdateCellDimensionsDisplay();
        RegeneratePreviewGrid();
    }

    // ──────────────────────────────────────────────────────────────
    // Helpers
    // ──────────────────────────────────────────────────────────────

    void MainWindow::UpdateCellDimensionsDisplay()
    {
        if (!m_isLoaded) return;
        auto imgWBox = NbImageW();
        auto imgHBox = NbImageH();
        auto txtCell = TxtCellDimensions();
        if (!imgWBox || !imgHBox || !txtCell) return;

        double imgW = imgWBox.Value();
        double imgH = imgHBox.Value();
        if (std::isnan(imgW) || std::isnan(imgH)) return;

        bool isCm = RadioCm() && RadioCm().IsChecked() && RadioCm().IsChecked().Value();
        hstring unit = isCm ? L"cm" : L"in";
        txtCell.Text(L"Cell: " + to_hstring(imgW) + L" \u00D7 " + to_hstring(imgH) + L" " + unit);
    }

    double MainWindow::GetPixelsPerUnit()
    {
        constexpr double dpi = 300.0;
        if (RadioCm())
        {
            auto c = RadioCm().IsChecked();
            if (c && c.Value()) return dpi / 2.54;
        }
        return dpi;
    }

    // ──────────────────────────────────────────────────────────────
    // Settings events
    // ──────────────────────────────────────────────────────────────

    void MainWindow::OnUnitChanged(IInspectable const&, RoutedEventArgs const&)
    {
        if (!m_isLoaded) return;

        bool toCm = RadioCm() && RadioCm().IsChecked() && RadioCm().IsChecked().Value();
        double factor = toCm ? 2.54 : (1.0 / 2.54);

        auto convert = [&](NumberBox const& box) {
            if (box) box.Value(box.Value() * factor);
        };

        m_isLoaded = false;
        convert(NbSheetW());
        convert(NbSheetH());
        convert(NbImageW());
        convert(NbImageH());
        convert(NbGap());
        m_isLoaded = true;

        UpdateSheetSize();
        RefitCropContainer();
        UpdateCellDimensionsDisplay();
        RegeneratePreviewGrid();
    }

    void MainWindow::OnSettingsChanged(NumberBox const&, NumberBoxValueChangedEventArgs const& args)
    {
        if (!m_isLoaded) return;
        if (std::isnan(args.NewValue())) return;
        UpdateSheetSize();
        RefitCropContainer();
        UpdateCellDimensionsDisplay();
        RegeneratePreviewGrid();
    }

    void MainWindow::UpdateSheetSize()
    {
        if (!m_isLoaded) return;
        auto grid = PreviewGrid();
        auto sw = NbSheetW();
        auto sh = NbSheetH();
        if (!grid || !sw || !sh) return;

        double w = sw.Value() * GetPixelsPerUnit();
        double h = sh.Value() * GetPixelsPerUnit();
        if (std::isnan(w) || std::isnan(h)) return;
        if (w < 1) w = 1;
        if (h < 1) h = 1;
        grid.Width(w);
        grid.Height(h);
    }

    // ──────────────────────────────────────────────────────────────
    // Crop-container sizing (match image cell aspect ratio)
    // ──────────────────────────────────────────────────────────────

    void MainWindow::RefitCropContainer()
    {
        auto container = CropContainer();
        auto parent = CropParentContainer();
        auto imgWBox = NbImageW();
        auto imgHBox = NbImageH();
        if (!container || !parent || !imgWBox || !imgHBox) return;

        double availW = parent.ActualWidth();
        double availH = parent.ActualHeight();
        if (availW < 10 || availH < 10) return;

        double imgW = imgWBox.Value();
        double imgH = imgHBox.Value();
        if (std::isnan(imgW) || std::isnan(imgH) || imgW <= 0 || imgH <= 0) return;

        double ratio = imgW / imgH;
        double candidateW = availW;
        double candidateH = availW / ratio;
        if (candidateH > availH) { candidateH = availH; candidateW = availH * ratio; }

        container.Width(std::max(1.0, candidateW - 4));
        container.Height(std::max(1.0, candidateH - 4));
    }

    void MainWindow::OnCropSizeChanged(IInspectable const&, SizeChangedEventArgs const&)
    {
        RefitCropContainer();
    }

    // ──────────────────────────────────────────────────────────────
    // OPTIMAL PLACEMENT ALGORITHM
    //
    // Layout model:  [gap] cell [gap] cell [gap] ... cell [gap]
    //
    // For dimension D, cell size s, gap g:
    //   Count N = floor( (D - g) / (s + g) )
    //   Cell i position = g + i * (s + g)
    //   Last cell right edge = N * (s + g)
    //
    // Zone boundary placed at N*(s+g) so that the next zone's
    // leading gap gives exactly one gap between cells.
    // ──────────────────────────────────────────────────────────────

    std::vector<ImagePlacement> MainWindow::CalculateOptimalPlacement(
        double sheetW, double sheetH, double imgW, double imgH, double gap)
    {
        double ppu  = GetPixelsPerUnit();
        double g    = gap * ppu;
        double sWpx = sheetW * ppu;
        double sHpx = sheetH * ppu;

        double cW_n = imgW * ppu;       // normal cell W
        double cH_n = imgH * ppu;       // normal cell H
        double cW_r = imgH * ppu;       // rotated cell W  (swapped)
        double cH_r = imgW * ppu;       // rotated cell H

        // Count: g + N*(s+g) <= D  =>  N = floor((D-g)/(s+g))
        auto fitCount = [](double D, double s, double g) -> int {
            if (s <= 0 || D < s + 2.0 * g) return 0;
            return static_cast<int>(std::floor((D - g) / (s + g)));
        };

        // Build uniform grid in sub-area [ox..ox+aW) x [oy..oy+aH)
        auto buildGrid = [&](double ox, double oy, double aW, double aH,
                             double cW, double cH, bool rot) -> std::vector<ImagePlacement>
        {
            int cols = fitCount(aW, cW, g);
            int rows = fitCount(aH, cH, g);
            std::vector<ImagePlacement> out;
            if (cols <= 0 || rows <= 0) return out;
            out.reserve(static_cast<size_t>(rows) * cols);
            for (int r = 0; r < rows; ++r)
                for (int c = 0; c < cols; ++c)
                    out.push_back({
                        ox + g + c * (cW + g),
                        oy + g + r * (cH + g),
                        cW, cH, rot });
            return out;
        };

        // Space consumed by N rows/cols (bottom/right edge of last cell).
        // Does NOT include trailing gap; the next zone supplies its own
        // leading gap, yielding exactly one gap at the boundary.
        auto bandH = [&](int n, double cH) -> double {
            return (n <= 0) ? 0.0 : n * (cH + g);
        };
        auto bandW = [&](int n, double cW) -> double {
            return (n <= 0) ? 0.0 : n * (cW + g);
        };

        int bestCount = 0;
        std::vector<ImagePlacement> best;
        hstring bestDesc;

        auto tryCandidate = [&](std::vector<ImagePlacement>& p, hstring desc) {
            int cnt = static_cast<int>(p.size());
            if (cnt > bestCount) {
                bestCount = cnt;
                best = std::move(p);
                bestDesc = desc;
            }
        };

        auto fmtGrid = [](int rows, int cols, bool rot) -> hstring {
            return to_hstring(rows) + L"\u00D7" + to_hstring(cols) + (rot ? L"R" : L"");
        };

        // Strategy 1: pure normal
        {
            int c = fitCount(sWpx, cW_n, g), r = fitCount(sHpx, cH_n, g);
            auto p = buildGrid(0, 0, sWpx, sHpx, cW_n, cH_n, false);
            tryCandidate(p, to_hstring((int)p.size()) + L" (" + fmtGrid(r, c, false) + L")");
        }

        // Strategy 2: pure rotated
        {
            int c = fitCount(sWpx, cW_r, g), r = fitCount(sHpx, cH_r, g);
            auto p = buildGrid(0, 0, sWpx, sHpx, cW_r, cH_r, true);
            tryCandidate(p, to_hstring((int)p.size()) + L" (" + fmtGrid(r, c, true) + L")");
        }

        bool isSquare = std::abs(imgW - imgH) < 0.001;
        if (!isSquare)
        {
            int maxNR = fitCount(sHpx, cH_n, g);
            int maxRR = fitCount(sHpx, cH_r, g);
            int maxNC = fitCount(sWpx, cW_n, g);
            int maxRC = fitCount(sWpx, cW_r, g);

            // Strategies 3-4: horizontal band split
            for (int nR = 0; nR <= maxNR; ++nR) {
                double sy = bandH(nR, cH_n);
                auto pT = buildGrid(0, 0, sWpx, sy, cW_n, cH_n, false);
                auto pB = buildGrid(0, sy, sWpx, sHpx - sy, cW_r, cH_r, true);
                pT.insert(pT.end(), pB.begin(), pB.end());
                tryCandidate(pT, to_hstring((int)pT.size()) + L" (H " + to_hstring(nR) + L"n+r)");
            }
            for (int nR = 0; nR <= maxRR; ++nR) {
                double sy = bandH(nR, cH_r);
                auto pT = buildGrid(0, 0, sWpx, sy, cW_r, cH_r, true);
                auto pB = buildGrid(0, sy, sWpx, sHpx - sy, cW_n, cH_n, false);
                pT.insert(pT.end(), pB.begin(), pB.end());
                tryCandidate(pT, to_hstring((int)pT.size()) + L" (H " + to_hstring(nR) + L"r+n)");
            }

            // Strategies 5-6: vertical band split
            for (int nC = 0; nC <= maxNC; ++nC) {
                double sx = bandW(nC, cW_n);
                auto pL = buildGrid(0, 0, sx, sHpx, cW_n, cH_n, false);
                auto pR = buildGrid(sx, 0, sWpx - sx, sHpx, cW_r, cH_r, true);
                pL.insert(pL.end(), pR.begin(), pR.end());
                tryCandidate(pL, to_hstring((int)pL.size()) + L" (V " + to_hstring(nC) + L"n+r)");
            }
            for (int nC = 0; nC <= maxRC; ++nC) {
                double sx = bandW(nC, cW_r);
                auto pL = buildGrid(0, 0, sx, sHpx, cW_r, cH_r, true);
                auto pR = buildGrid(sx, 0, sWpx - sx, sHpx, cW_n, cH_n, false);
                pL.insert(pL.end(), pR.begin(), pR.end());
                tryCandidate(pL, to_hstring((int)pL.size()) + L" (V " + to_hstring(nC) + L"r+n)");
            }

            // Strategies 7-8: L-shaped 3-zone (top full-width, bottom split)
            for (int topR = 0; topR <= maxNR; ++topR) {
                double sy = bandH(topR, cH_n);
                double bH = sHpx - sy;
                if (bH < g) continue;

                for (int blc = 0, mx = fitCount(sWpx, cW_r, g); blc <= mx; ++blc) {
                    double sx = bandW(blc, cW_r);
                    auto pT  = buildGrid(0, 0, sWpx, sy, cW_n, cH_n, false);
                    auto pBL = buildGrid(0, sy, sx, bH, cW_r, cH_r, true);
                    auto pBR = buildGrid(sx, sy, sWpx - sx, bH, cW_n, cH_n, false);
                    pT.insert(pT.end(), pBL.begin(), pBL.end());
                    pT.insert(pT.end(), pBR.begin(), pBR.end());
                    tryCandidate(pT, to_hstring((int)pT.size()) + L" (L n|r+n)");
                }
                for (int blc = 0, mx = fitCount(sWpx, cW_n, g); blc <= mx; ++blc) {
                    double sx = bandW(blc, cW_n);
                    auto pT  = buildGrid(0, 0, sWpx, sy, cW_n, cH_n, false);
                    auto pBL = buildGrid(0, sy, sx, bH, cW_n, cH_n, false);
                    auto pBR = buildGrid(sx, sy, sWpx - sx, bH, cW_r, cH_r, true);
                    pT.insert(pT.end(), pBL.begin(), pBL.end());
                    pT.insert(pT.end(), pBR.begin(), pBR.end());
                    tryCandidate(pT, to_hstring((int)pT.size()) + L" (L n|n+r)");
                }
            }
            for (int topR = 0; topR <= maxRR; ++topR) {
                double sy = bandH(topR, cH_r);
                double bH = sHpx - sy;
                if (bH < g) continue;

                for (int blc = 0, mx = fitCount(sWpx, cW_n, g); blc <= mx; ++blc) {
                    double sx = bandW(blc, cW_n);
                    auto pT  = buildGrid(0, 0, sWpx, sy, cW_r, cH_r, true);
                    auto pBL = buildGrid(0, sy, sx, bH, cW_n, cH_n, false);
                    auto pBR = buildGrid(sx, sy, sWpx - sx, bH, cW_r, cH_r, true);
                    pT.insert(pT.end(), pBL.begin(), pBL.end());
                    pT.insert(pT.end(), pBR.begin(), pBR.end());
                    tryCandidate(pT, to_hstring((int)pT.size()) + L" (L r|n+r)");
                }
                for (int blc = 0, mx = fitCount(sWpx, cW_r, g); blc <= mx; ++blc) {
                    double sx = bandW(blc, cW_r);
                    auto pT  = buildGrid(0, 0, sWpx, sy, cW_r, cH_r, true);
                    auto pBL = buildGrid(0, sy, sx, bH, cW_r, cH_r, true);
                    auto pBR = buildGrid(sx, sy, sWpx - sx, bH, cW_n, cH_n, false);
                    pT.insert(pT.end(), pBL.begin(), pBL.end());
                    pT.insert(pT.end(), pBR.begin(), pBR.end());
                    tryCandidate(pT, to_hstring((int)pT.size()) + L" (L r|r+n)");
                }
            }
        }

        Log(L"Optimal: " + to_hstring(bestCount) + L" \u2013 " + bestDesc);

        if (TxtLayoutInfo())
            TxtLayoutInfo().Text(bestCount > 0 ? bestDesc : L"Images don't fit on sheet");

        return best;
    }

    // ──────────────────────────────────────────────────────────────
    // Preview grid rendering
    // ──────────────────────────────────────────────────────────────

    winrt::Windows::Foundation::IAsyncAction MainWindow::RegeneratePreviewGrid()
    {
        if (!m_isLoaded) co_return;
        auto strong = get_strong();

        auto grid = PreviewGrid();
        if (!grid) co_return;

        auto swBox   = NbSheetW();
        auto shBox   = NbSheetH();
        auto imgWBox = NbImageW();
        auto imgHBox = NbImageH();
        auto gapBox  = NbGap();
        if (!swBox || !shBox || !imgWBox || !imgHBox || !gapBox) co_return;

        // Clear
        grid.Children().Clear();
        grid.RowDefinitions().Clear();
        grid.ColumnDefinitions().Clear();
        grid.RowSpacing(0);
        grid.ColumnSpacing(0);
        grid.Padding({ 0, 0, 0, 0 });
        m_outlineBorders.clear();

        double sheetW = swBox.Value();
        double sheetH = shBox.Value();
        double imgW   = imgWBox.Value();
        double imgH   = imgHBox.Value();
        double gap    = gapBox.Value();

        if (std::isnan(sheetW) || std::isnan(sheetH) || std::isnan(imgW) ||
            std::isnan(imgH)   || std::isnan(gap)) co_return;
        if (sheetW <= 0 || sheetH <= 0 || imgW <= 0 || imgH <= 0 || gap < 0) co_return;

        m_currentPlacements = CalculateOptimalPlacement(sheetW, sheetH, imgW, imgH, gap);

        // Request #1: blank sheet until user applies a crop
        if (!m_croppedStamp || m_currentPlacements.empty()) co_return;

        for (auto& p : m_currentPlacements)
        {
            auto& bmp = p.rotated ? m_croppedStampRotated : m_croppedStamp;
            if (!bmp) continue;

            Image img;
            SoftwareBitmapSource src;
            try { co_await src.SetBitmapAsync(bmp); }
            catch (hresult_error const& ex) {
                Log(L"SetBitmapAsync failed: " + ex.message());
                continue;
            }

            img.Source(src);
            img.Stretch(Stretch::Uniform);

            // Request #2: red outline on preview (hidden during export)
            Border outline;
            outline.BorderBrush(SolidColorBrush(
                Microsoft::UI::ColorHelper::FromArgb(255, 220, 50, 50)));
            outline.BorderThickness({ 1, 1, 1, 1 });
            outline.Width(p.w);
            outline.Height(p.h);
            outline.HorizontalAlignment(HorizontalAlignment::Left);
            outline.VerticalAlignment(VerticalAlignment::Top);
            outline.Margin({ p.x, p.y, 0, 0 });
            outline.Child(img);

            grid.Children().Append(outline);
            m_outlineBorders.push_back(outline);
        }
    }

    // ──────────────────────────────────────────────────────────────
    // Outline visibility for clean export
    // ──────────────────────────────────────────────────────────────

    void MainWindow::SetOutlinesVisible(bool visible)
    {
        auto brush = visible
            ? SolidColorBrush(Microsoft::UI::ColorHelper::FromArgb(255, 220, 50, 50))
            : SolidColorBrush(Microsoft::UI::Colors::Transparent());
        for (auto& b : m_outlineBorders)
            if (b) b.BorderBrush(brush);
    }

    // ──────────────────────────────────────────────────────────────
    // Crop / Apply / Rotate
    // ──────────────────────────────────────────────────────────────

    winrt::fire_and_forget MainWindow::BtnApplyCrop_Click(IInspectable const&, RoutedEventArgs const&)
    {
        if (!m_originalBitmap) co_return;
        auto strong = get_strong();

        auto stamp = co_await CaptureCropAsBitmap();
        if (!stamp) co_return;

        m_croppedStamp = stamp;
        m_croppedStampRotated = co_await RotateBitmap90(stamp);
        co_await RegeneratePreviewGrid();
    }

    winrt::Windows::Foundation::IAsyncOperation<SoftwareBitmap> MainWindow::CaptureCropAsBitmap()
    {
        if (!m_originalBitmap) co_return nullptr;
        auto scroller = CropScrollViewer();
        if (!scroller) co_return nullptr;

        double srcW = static_cast<double>(m_originalBitmap.PixelWidth());
        double srcH = static_cast<double>(m_originalBitmap.PixelHeight());

        float  zoom  = scroller.ZoomFactor();
        double hOff  = scroller.HorizontalOffset();
        double vOff  = scroller.VerticalOffset();
        double viewW = scroller.ViewportWidth();
        double viewH = scroller.ViewportHeight();

        double cropX = hOff / zoom;
        double cropY = vOff / zoom;
        double cropW = viewW / zoom;
        double cropH = viewH / zoom;

        cropX = std::max(0.0, cropX);
        cropY = std::max(0.0, cropY);
        cropW = std::min(cropW, srcW - cropX);
        cropH = std::min(cropH, srcH - cropY);
        if (cropW <= 0 || cropH <= 0) co_return nullptr;

        BitmapBounds bounds;
        bounds.X      = static_cast<uint32_t>(std::floor(cropX));
        bounds.Y      = static_cast<uint32_t>(std::floor(cropY));
        bounds.Width   = std::max(1u, static_cast<uint32_t>(std::floor(cropW)));
        bounds.Height  = std::max(1u, static_cast<uint32_t>(std::floor(cropH)));

        if (bounds.X + bounds.Width  > static_cast<uint32_t>(srcW))
            bounds.Width  = static_cast<uint32_t>(srcW) - bounds.X;
        if (bounds.Y + bounds.Height > static_cast<uint32_t>(srcH))
            bounds.Height = static_cast<uint32_t>(srcH) - bounds.Y;

        try
        {
            InMemoryRandomAccessStream tmp;
            auto enc = co_await BitmapEncoder::CreateAsync(BitmapEncoder::PngEncoderId(), tmp);
            enc.SetSoftwareBitmap(m_originalBitmap);
            enc.BitmapTransform().Bounds(bounds);
            co_await enc.FlushAsync();

            tmp.Seek(0);
            auto dec = co_await BitmapDecoder::CreateAsync(tmp);
            auto bmp = co_await dec.GetSoftwareBitmapAsync();
            if (bmp.BitmapPixelFormat() != BitmapPixelFormat::Bgra8 ||
                bmp.BitmapAlphaMode()   != BitmapAlphaMode::Premultiplied)
                bmp = SoftwareBitmap::Convert(bmp, BitmapPixelFormat::Bgra8, BitmapAlphaMode::Premultiplied);
            co_return bmp;
        }
        catch (hresult_error const& ex) {
            Log(L"Crop failed: " + ex.message());
            co_return nullptr;
        }
    }

    winrt::Windows::Foundation::IAsyncOperation<SoftwareBitmap> MainWindow::RotateBitmap90(SoftwareBitmap bmp)
    {
        if (!bmp) co_return nullptr;
        try
        {
            InMemoryRandomAccessStream stream;
            auto enc = co_await BitmapEncoder::CreateAsync(BitmapEncoder::PngEncoderId(), stream);
            enc.SetSoftwareBitmap(bmp);
            enc.BitmapTransform().Rotation(BitmapRotation::Clockwise90Degrees);
            co_await enc.FlushAsync();
            stream.Seek(0);
            auto dec = co_await BitmapDecoder::CreateAsync(stream);
            auto rotated = co_await dec.GetSoftwareBitmapAsync();
            if (rotated.BitmapPixelFormat() != BitmapPixelFormat::Bgra8 ||
                rotated.BitmapAlphaMode()   != BitmapAlphaMode::Premultiplied)
                rotated = SoftwareBitmap::Convert(rotated, BitmapPixelFormat::Bgra8, BitmapAlphaMode::Premultiplied);
            co_return rotated;
        }
        catch (hresult_error const& ex) {
            Log(L"RotateBitmap90 failed: " + ex.message());
            co_return nullptr;
        }
    }

    winrt::fire_and_forget MainWindow::BtnRotate_Click(IInspectable const&, RoutedEventArgs const&)
    {
        if (!m_originalBitmap) co_return;
        auto strong = get_strong();

        try
        {
            auto rotated = co_await RotateBitmap90(m_originalBitmap);
            if (!rotated) co_return;

            m_originalBitmap = rotated;
            SoftwareBitmapSource src;
            co_await src.SetBitmapAsync(m_originalBitmap);
            if (SourceImageControl()) SourceImageControl().Source(src);

            ZoomToFit();

            m_croppedStamp = nullptr;
            m_croppedStampRotated = nullptr;
            co_await RegeneratePreviewGrid();
        }
        catch (hresult_error const& ex) {
            Log(L"Rotation exception: " + ex.message());
        }
    }

    // ──────────────────────────────────────────────────────────────
    // Zoom-to-fit
    // ──────────────────────────────────────────────────────────────

    void MainWindow::ZoomToFit()
    {
        try
        {
            auto scroller = CropScrollViewer();
            if (!scroller || !m_originalBitmap) return;

            double imgW = static_cast<double>(m_originalBitmap.PixelWidth());
            double imgH = static_cast<double>(m_originalBitmap.PixelHeight());
            if (imgW <= 0 || imgH <= 0) return;

            double vpW = scroller.ViewportWidth();
            double vpH = scroller.ViewportHeight();
            if (vpW <= 0 || vpH <= 0) {
                auto c = CropContainer();
                if (c) { vpW = c.ActualWidth(); vpH = c.ActualHeight(); }
                if (vpW <= 0 || vpH <= 0) return;
            }

            float z = static_cast<float>(std::min(vpW / imgW, vpH / imgH));
            z = std::clamp(z, 0.1f, 10.0f);

            m_zoomingFromMouse = true;
            scroller.ChangeView(0.0, 0.0, z, true);
            if (ZoomSlider()) ZoomSlider().Value(z);
            m_zoomingFromMouse = false;
        }
        catch (hresult_error const&) {
            m_zoomingFromMouse = false;
        }
    }

    // ──────────────────────────────────────────────────────────────
    // Save sheet
    // ──────────────────────────────────────────────────────────────

    winrt::fire_and_forget MainWindow::BtnSaveSheet_Click(IInspectable const&, RoutedEventArgs const&)
    {
        auto grid = PreviewGrid();
        if (!grid || m_currentPlacements.empty() || !m_croppedStamp) co_return;
        auto strong = get_strong();

        FileSavePicker picker;
        auto initWnd{ picker.as<::IInitializeWithWindow>() };
        HWND hwnd;
        auto windowNative{ this->try_as<::IWindowNative>() };
        windowNative->get_WindowHandle(&hwnd);
        initWnd->Initialize(hwnd);
        picker.FileTypeChoices().Insert(L"JPEG", single_threaded_vector<hstring>({ L".jpg" }));
        picker.FileTypeChoices().Insert(L"PNG",  single_threaded_vector<hstring>({ L".png" }));
        picker.SuggestedFileName(L"PassportSheet");

        StorageFile file = co_await picker.PickSaveFileAsync();
        if (!file) co_return;

        try
        {
            // Hide outlines for clean export
            SetOutlinesVisible(false);

            int targetW = static_cast<int>(grid.Width());
            int targetH = static_cast<int>(grid.Height());

            constexpr int kMaxDim = 16384;
            double scale = 1.0;
            if (targetW > kMaxDim || targetH > kMaxDim) {
                scale = std::min(static_cast<double>(kMaxDim) / targetW,
                                 static_cast<double>(kMaxDim) / targetH);
                targetW = static_cast<int>(targetW * scale);
                targetH = static_cast<int>(targetH * scale);
            }

            RenderTargetBitmap rtb;
            co_await rtb.RenderAsync(grid, targetW, targetH);

            SoftwareBitmap sb(BitmapPixelFormat::Bgra8,
                              rtb.PixelWidth(), rtb.PixelHeight(),
                              BitmapAlphaMode::Premultiplied);
            auto buffer = co_await rtb.GetPixelsAsync();
            sb.CopyFromBuffer(buffer);

            SetOutlinesVisible(true);

            auto ext = file.FileType();
            auto encoderId = (ext == L".png") ? BitmapEncoder::PngEncoderId()
                                              : BitmapEncoder::JpegEncoderId();

            auto stream = co_await file.OpenAsync(FileAccessMode::ReadWrite);
            auto enc = co_await BitmapEncoder::CreateAsync(encoderId, stream);
            enc.SetSoftwareBitmap(sb);
            enc.BitmapTransform().InterpolationMode(BitmapInterpolationMode::Fant);
            co_await enc.FlushAsync();
        }
        catch (hresult_error const& ex) {
            SetOutlinesVisible(true);
            Log(L"Save failed: " + ex.message());
        }
    }

    // ──────────────────────────────────────────────────────────────
    // Image loading
    // ──────────────────────────────────────────────────────────────

    winrt::fire_and_forget MainWindow::BtnPickImage_Click(IInspectable const&, RoutedEventArgs const&)
    {
        auto strong = get_strong();

        FileOpenPicker picker;
        auto initWnd{ picker.as<::IInitializeWithWindow>() };
        HWND hwnd;
        auto windowNative{ this->try_as<::IWindowNative>() };
        windowNative->get_WindowHandle(&hwnd);
        initWnd->Initialize(hwnd);
        picker.FileTypeFilter().ReplaceAll({ L".jpg", L".jpeg", L".png", L".bmp", L".tiff" });
        StorageFile file = co_await picker.PickSingleFileAsync();
        if (file) co_await LoadImageFromFile(file);
    }

    winrt::Windows::Foundation::IAsyncAction MainWindow::LoadImageFromFile(StorageFile file)
    {
        auto strong = get_strong();

        try
        {
            auto stream = co_await file.OpenAsync(FileAccessMode::Read);
            auto decoder = co_await BitmapDecoder::CreateAsync(stream);
            m_originalBitmap = co_await decoder.GetSoftwareBitmapAsync();

            if (m_originalBitmap.BitmapPixelFormat() != BitmapPixelFormat::Bgra8 ||
                m_originalBitmap.BitmapAlphaMode()   != BitmapAlphaMode::Premultiplied)
                m_originalBitmap = SoftwareBitmap::Convert(
                    m_originalBitmap, BitmapPixelFormat::Bgra8, BitmapAlphaMode::Premultiplied);

            SoftwareBitmapSource src;
            co_await src.SetBitmapAsync(m_originalBitmap);
            if (SourceImageControl()) SourceImageControl().Source(src);

            m_croppedStamp = nullptr;
            m_croppedStampRotated = nullptr;

            co_await RegeneratePreviewGrid();

            // Zoom-to-fit after layout settles
            auto weak = get_weak();
            if (auto dq = this->DispatcherQueue())
                dq.TryEnqueue([weak]() { if (auto s = weak.get()) s->ZoomToFit(); });
        }
        catch (hresult_error const& ex) {
            Log(L"LoadImageFromFile failed: " + ex.message());
        }
    }

    // ──────────────────────────────────────────────────────────────
    // Pointer: pan & zoom
    // ──────────────────────────────────────────────────────────────

    void MainWindow::OnImagePointerPressed(IInspectable const& sender, PointerRoutedEventArgs const& e)
    {
        auto el = sender.try_as<UIElement>();
        if (!el) return;
        el.CapturePointer(e.Pointer());
        m_isDragging = true;
        if (CropScrollViewer())
            m_lastPoint = e.GetCurrentPoint(CropScrollViewer()).Position();
        e.Handled(true);
    }

    void MainWindow::OnImagePointerMoved(IInspectable const&, PointerRoutedEventArgs const& e)
    {
        if (!m_isDragging) return;
        auto scroller = CropScrollViewer();
        if (!scroller) return;

        auto pt = e.GetCurrentPoint(scroller).Position();
        double dx = pt.X - m_lastPoint.X;
        double dy = pt.Y - m_lastPoint.Y;

        // disableAnimation = true  =>  instant, lag-free panning
        scroller.ChangeView(
            scroller.HorizontalOffset() - dx,
            scroller.VerticalOffset()   - dy,
            nullptr, true);

        m_lastPoint = pt;
        e.Handled(true);
    }

    void MainWindow::OnImagePointerReleased(IInspectable const& sender, PointerRoutedEventArgs const& e)
    {
        m_isDragging = false;
        auto el = sender.try_as<UIElement>();
        if (el) el.ReleasePointerCapture(e.Pointer());
        e.Handled(true);
    }

    void MainWindow::OnImagePointerWheelChanged(IInspectable const&, PointerRoutedEventArgs const& e)
    {
        auto scroller = CropScrollViewer();
        if (!scroller) return;

        int delta = e.GetCurrentPoint(scroller).Properties().MouseWheelDelta();
        double oldZoom = scroller.ZoomFactor();
        double step = oldZoom * 0.15;
        double newZoom = (delta > 0) ? oldZoom + step : oldZoom - step;
        newZoom = std::clamp(newZoom, 0.1, 10.0);
        if (std::abs(newZoom - oldZoom) < 1e-6) { e.Handled(true); return; }

        auto pt = e.GetCurrentPoint(scroller).Position();
        double hOff = scroller.HorizontalOffset();
        double vOff = scroller.VerticalOffset();
        double newH = ((hOff + pt.X) / oldZoom) * newZoom - pt.X;
        double newV = ((vOff + pt.Y) / oldZoom) * newZoom - pt.Y;

        m_zoomingFromMouse = true;
        if (ZoomSlider()) ZoomSlider().Value(newZoom);
        scroller.ChangeView(newH, newV, static_cast<float>(newZoom), true);
        m_zoomingFromMouse = false;
        e.Handled(true);
    }

    // ──────────────────────────────────────────────────────────────
    // Drag & drop
    // ──────────────────────────────────────────────────────────────

    void MainWindow::OnDragOver(IInspectable const&, DragEventArgs const& e)
    {
        if (e.DataView().Contains(StandardDataFormats::StorageItems()))
            e.AcceptedOperation(DataPackageOperation::Copy);
    }

    winrt::fire_and_forget MainWindow::OnDrop(IInspectable const&, DragEventArgs const& e)
    {
        auto strong = get_strong();
        if (!e.DataView().Contains(StandardDataFormats::StorageItems())) co_return;
        auto items = co_await e.DataView().GetStorageItemsAsync();
        if (items.Size() > 0 && items.GetAt(0).IsOfType(StorageItemTypes::File))
            co_await LoadImageFromFile(items.GetAt(0).as<StorageFile>());
    }
}
