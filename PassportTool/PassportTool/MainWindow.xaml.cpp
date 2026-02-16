#include "pch.h"
#include "MainWindow.xaml.h"
#if __has_include("MainWindow.g.cpp")
#include "MainWindow.g.cpp"
#endif
#include <cmath>
#include <algorithm>
#include <vector>
#include <numeric>

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
        Log(L"MainWindow Constructor started.");

        // Number-box formatting (2 decimal places)
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

        // Slider → ScrollViewer zoom
        if (ZoomSlider())
        {
            ZoomSlider().ValueChanged([this](auto&&, auto&& args)
            {
                if (m_zoomingFromMouse) return;
                m_zoomingFromSlider = true;
                if (CropScrollViewer())
                    CropScrollViewer().ChangeView(nullptr, nullptr, static_cast<float>(args.NewValue()));
                m_zoomingFromSlider = false;
            });
        }

        // ScrollViewer → Slider sync
        if (CropScrollViewer())
        {
            CropScrollViewer().ViewChanged([this](auto&&, auto&&)
            {
                if (m_zoomingFromSlider || m_zoomingFromMouse) return;
                if (CropScrollViewer() && ZoomSlider())
                    ZoomSlider().Value(CropScrollViewer().ZoomFactor());
            });
        }

        Log(L"MainWindow Constructor completed.");
    }

    void MainWindow::Log(winrt::hstring const& message)
    {
        std::wstring msg = L"[PassportTool] " + std::wstring(message) + L"\r\n";
        OutputDebugStringW(msg.c_str());
    }

    void MainWindow::OnWindowLoaded(IInspectable const&, RoutedEventArgs const&)
    {
        Log(L"Window Loaded Event.");
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
        bool isCm = RadioCm() && RadioCm().IsChecked() && RadioCm().IsChecked().Value();
        hstring unit = isCm ? L"cm" : L"in";
        txtCell.Text(L"Cell: " + to_hstring(imgW) + L" \u00D7 " + to_hstring(imgH) + L" " + unit);
    }

    double MainWindow::GetPixelsPerUnit()
    {
        double dpi = 300.0;
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

        m_isLoaded = false;           // suppress cascading events
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

    void MainWindow::OnSettingsChanged(NumberBox const& sender, NumberBoxValueChangedEventArgs const& args)
    {
        if (!m_isLoaded) return;
        // NumberBox produces NaN when the field is cleared - ignore it
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
    // Crop-container sizing (maintains aspect ratio of image cell)
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
        if (imgW <= 0 || imgH <= 0) return;

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
    // Tries pure-normal, pure-rotated, and horizontal / vertical
    // band splits (both orderings) to maximise image count.
    // For square images rotation is a no-op so the pure grid wins.
    // ──────────────────────────────────────────────────────────────

    std::vector<ImagePlacement> MainWindow::CalculateOptimalPlacement(
        double sheetW, double sheetH, double imgW, double imgH, double gap)
    {
        double ppu = GetPixelsPerUnit();
        double gapPx = gap * ppu;
        double sheetWPx = sheetW * ppu;
        double sheetHPx = sheetH * ppu;

        // Normal cell (image as-is)
        double cW_n = imgW * ppu;
        double cH_n = imgH * ppu;

        // Rotated cell (swap dimensions)
        double cW_r = imgH * ppu;
        double cH_r = imgW * ppu;

        // How many items of size `s` fit in dimension `D` with gap `g`?
        auto calcCount = [](double D, double s, double g) -> int {
            if (D < g || s <= 0) return 0;
            return static_cast<int>(std::floor((D - g) / (s + g)));
        };

        // Build a vector of placements for a grid block at (ox,oy) in area (aW x aH)
        auto buildGrid = [&](double ox, double oy, double aW, double aH,
                             double cW, double cH, double g, bool rot) -> std::vector<ImagePlacement>
        {
            int cols = calcCount(aW, cW, g);
            int rows = calcCount(aH, cH, g);
            std::vector<ImagePlacement> out;
            out.reserve(rows * cols);
            for (int r = 0; r < rows; ++r)
                for (int c = 0; c < cols; ++c)
                    out.push_back({ ox + g + c * (cW + g),
                                    oy + g + r * (cH + g),
                                    cW, cH, rot });
            return out;
        };

        // Height consumed by `n` rows of cell-height `cH` with gap
        auto bandH = [&](int n, double cH) -> double {
            if (n <= 0) return 0;
            return gapPx + n * (cH + gapPx);
        };
        auto bandW = [&](int n, double cW) -> double {
            if (n <= 0) return 0;
            return gapPx + n * (cW + gapPx);
        };

        int bestCount = 0;
        std::vector<ImagePlacement> bestPlacements;
        hstring bestDesc;

        auto tryCandidate = [&](std::vector<ImagePlacement>& placements, hstring desc) {
            int cnt = static_cast<int>(placements.size());
            if (cnt > bestCount) {
                bestCount = cnt;
                bestPlacements = placements;
                bestDesc = desc;
            }
        };

        // ── Strategy 1: pure normal ──
        {
            auto p = buildGrid(0, 0, sheetWPx, sheetHPx, cW_n, cH_n, gapPx, false);
            int cols = calcCount(sheetWPx, cW_n, gapPx);
            int rows = calcCount(sheetHPx, cH_n, gapPx);
            tryCandidate(p, to_hstring((int)p.size()) + L" (" + to_hstring(rows) + L"\u00D7" + to_hstring(cols) + L")");
        }

        // ── Strategy 2: pure rotated ──
        {
            auto p = buildGrid(0, 0, sheetWPx, sheetHPx, cW_r, cH_r, gapPx, true);
            int cols = calcCount(sheetWPx, cW_r, gapPx);
            int rows = calcCount(sheetHPx, cH_r, gapPx);
            tryCandidate(p, to_hstring((int)p.size()) + L" (" + to_hstring(rows) + L"\u00D7" + to_hstring(cols) + L") Rot");
        }

        // Skip mixed strategies for square images (rotation == identity)
        bool isSquare = std::abs(imgW - imgH) < 0.001;
        if (!isSquare)
        {
            int maxNormRows = calcCount(sheetHPx, cH_n, gapPx);
            int maxRotRows  = calcCount(sheetHPx, cH_r, gapPx);
            int maxNormCols = calcCount(sheetWPx, cW_n, gapPx);
            int maxRotCols  = calcCount(sheetWPx, cW_r, gapPx);

            // ── Strategy 3: H-split, normal top + rotated bottom ──
            for (int nR = 0; nR <= maxNormRows; ++nR)
            {
                double topH = bandH(nR, cH_n);
                double botH = sheetHPx - topH;
                auto pTop = buildGrid(0, 0, sheetWPx, topH, cW_n, cH_n, gapPx, false);
                auto pBot = buildGrid(0, topH, sheetWPx, botH, cW_r, cH_r, gapPx, true);
                pTop.insert(pTop.end(), pBot.begin(), pBot.end());
                tryCandidate(pTop, to_hstring((int)pTop.size()) + L" (H-split N+" + to_hstring(nR) + L"r)");
            }

            // ── Strategy 4: H-split, rotated top + normal bottom ──
            for (int nR = 0; nR <= maxRotRows; ++nR)
            {
                double topH = bandH(nR, cH_r);
                double botH = sheetHPx - topH;
                auto pTop = buildGrid(0, 0, sheetWPx, topH, cW_r, cH_r, gapPx, true);
                auto pBot = buildGrid(0, topH, sheetWPx, botH, cW_n, cH_n, gapPx, false);
                pTop.insert(pTop.end(), pBot.begin(), pBot.end());
                tryCandidate(pTop, to_hstring((int)pTop.size()) + L" (H-split R+" + to_hstring(nR) + L"n)");
            }

            // ── Strategy 5: V-split, normal left + rotated right ──
            for (int nC = 0; nC <= maxNormCols; ++nC)
            {
                double leftW = bandW(nC, cW_n);
                double rightW = sheetWPx - leftW;
                auto pL = buildGrid(0, 0, leftW, sheetHPx, cW_n, cH_n, gapPx, false);
                auto pR = buildGrid(leftW, 0, rightW, sheetHPx, cW_r, cH_r, gapPx, true);
                pL.insert(pL.end(), pR.begin(), pR.end());
                tryCandidate(pL, to_hstring((int)pL.size()) + L" (V-split N+" + to_hstring(nC) + L"r)");
            }

            // ── Strategy 6: V-split, rotated left + normal right ──
            for (int nC = 0; nC <= maxRotCols; ++nC)
            {
                double leftW = bandW(nC, cW_r);
                double rightW = sheetWPx - leftW;
                auto pL = buildGrid(0, 0, leftW, sheetHPx, cW_r, cH_r, gapPx, true);
                auto pR = buildGrid(leftW, 0, rightW, sheetHPx, cW_n, cH_n, gapPx, false);
                pL.insert(pL.end(), pR.begin(), pR.end());
                tryCandidate(pL, to_hstring((int)pL.size()) + L" (V-split R+" + to_hstring(nC) + L"n)");
            }

            // ── Strategy 7: L-shaped / 4-zone packing ──
            // Normal rows on top spanning full width, then in remaining height:
            //   rotated block on left, plus normal block in remaining rect on right
            for (int topRows = 0; topRows <= maxNormRows; ++topRows)
            {
                double topH = bandH(topRows, cH_n);
                double botH = sheetHPx - topH;

                // Bottom-left: rotated
                int botRotCols = calcCount(sheetWPx, cW_r, gapPx);
                for (int brc = 0; brc <= botRotCols; ++brc)
                {
                    double blW = bandW(brc, cW_r);
                    double brW = sheetWPx - blW;

                    auto pTop = buildGrid(0, 0, sheetWPx, topH, cW_n, cH_n, gapPx, false);
                    auto pBL  = buildGrid(0, topH, blW, botH, cW_r, cH_r, gapPx, true);
                    auto pBR  = buildGrid(blW, topH, brW, botH, cW_n, cH_n, gapPx, false);

                    pTop.insert(pTop.end(), pBL.begin(), pBL.end());
                    pTop.insert(pTop.end(), pBR.begin(), pBR.end());
                    tryCandidate(pTop, to_hstring((int)pTop.size()) + L" (L-shape)");
                }

                // Bottom-left: normal, bottom-right: rotated
                int botNormCols = calcCount(sheetWPx, cW_n, gapPx);
                for (int bnc = 0; bnc <= botNormCols; ++bnc)
                {
                    double blW = bandW(bnc, cW_n);
                    double brW = sheetWPx - blW;

                    auto pTop = buildGrid(0, 0, sheetWPx, topH, cW_n, cH_n, gapPx, false);
                    auto pBL  = buildGrid(0, topH, blW, botH, cW_n, cH_n, gapPx, false);
                    auto pBR  = buildGrid(blW, topH, brW, botH, cW_r, cH_r, gapPx, true);

                    pTop.insert(pTop.end(), pBL.begin(), pBL.end());
                    pTop.insert(pTop.end(), pBR.begin(), pBR.end());
                    tryCandidate(pTop, to_hstring((int)pTop.size()) + L" (L-shape2)");
                }
            }

            // ── Strategy 8: Rotated rows on top, then 4-zone below ──
            for (int topRows = 0; topRows <= maxRotRows; ++topRows)
            {
                double topH = bandH(topRows, cH_r);
                double botH = sheetHPx - topH;

                int botNormCols = calcCount(sheetWPx, cW_n, gapPx);
                for (int bnc = 0; bnc <= botNormCols; ++bnc)
                {
                    double blW = bandW(bnc, cW_n);
                    double brW = sheetWPx - blW;

                    auto pTop = buildGrid(0, 0, sheetWPx, topH, cW_r, cH_r, gapPx, true);
                    auto pBL  = buildGrid(0, topH, blW, botH, cW_n, cH_n, gapPx, false);
                    auto pBR  = buildGrid(blW, topH, brW, botH, cW_r, cH_r, gapPx, true);

                    pTop.insert(pTop.end(), pBL.begin(), pBL.end());
                    pTop.insert(pTop.end(), pBR.begin(), pBR.end());
                    tryCandidate(pTop, to_hstring((int)pTop.size()) + L" (L-shape3)");
                }

                int botRotCols = calcCount(sheetWPx, cW_r, gapPx);
                for (int brc = 0; brc <= botRotCols; ++brc)
                {
                    double blW = bandW(brc, cW_r);
                    double brW = sheetWPx - blW;

                    auto pTop = buildGrid(0, 0, sheetWPx, topH, cW_r, cH_r, gapPx, true);
                    auto pBL  = buildGrid(0, topH, blW, botH, cW_r, cH_r, gapPx, true);
                    auto pBR  = buildGrid(blW, topH, brW, botH, cW_n, cH_n, gapPx, false);

                    pTop.insert(pTop.end(), pBL.begin(), pBL.end());
                    pTop.insert(pTop.end(), pBR.begin(), pBR.end());
                    tryCandidate(pTop, to_hstring((int)pTop.size()) + L" (L-shape4)");
                }
            }
        }

        Log(L"Optimal placement: " + to_hstring(bestCount) + L" images – " + bestDesc);

        // Update layout info text
        if (TxtLayoutInfo())
        {
            if (bestCount > 0)
                TxtLayoutInfo().Text(bestDesc);
            else
                TxtLayoutInfo().Text(L"Images don't fit on sheet");
        }

        return bestPlacements;
    }

    // ──────────────────────────────────────────────────────────────
    // Preview grid rendering
    // ──────────────────────────────────────────────────────────────

    winrt::Windows::Foundation::IAsyncAction MainWindow::RegeneratePreviewGrid()
    {
        if (!m_isLoaded) co_return;
        auto strong = get_strong();
        Log(L"Regenerating Preview Grid...");

        auto grid = PreviewGrid();
        if (!grid) co_return;

        auto swBox   = NbSheetW();
        auto shBox   = NbSheetH();
        auto imgWBox = NbImageW();
        auto imgHBox = NbImageH();
        auto gapBox  = NbGap();
        if (!swBox || !shBox || !imgWBox || !imgHBox || !gapBox) co_return;

        // Clear previous content
        grid.Children().Clear();
        grid.RowDefinitions().Clear();
        grid.ColumnDefinitions().Clear();
        grid.RowSpacing(0);
        grid.ColumnSpacing(0);
        grid.Padding({ 0,0,0,0 });

        double sheetW = swBox.Value();
        double sheetH = shBox.Value();
        double imgW   = imgWBox.Value();
        double imgH   = imgHBox.Value();
        double gap    = gapBox.Value();

        // Guard against NaN from empty NumberBox
        if (std::isnan(sheetW) || std::isnan(sheetH) || std::isnan(imgW) || std::isnan(imgH) || std::isnan(gap))
            co_return;
        if (sheetW <= 0 || sheetH <= 0 || imgW <= 0 || imgH <= 0) co_return;

        // Compute optimal layout
        m_currentPlacements = CalculateOptimalPlacement(sheetW, sheetH, imgW, imgH, gap);

        if (m_currentPlacements.empty()) co_return;

        double ppu = GetPixelsPerUnit();
        double realImgWPx = imgW * ppu;
        double realImgHPx = imgH * ppu;

        for (auto& p : m_currentPlacements)
        {
            if (m_croppedStamp)
            {
                // Choose the right bitmap (normal or pre-rotated)
                auto& bmp = p.rotated ? m_croppedStampRotated : m_croppedStamp;
                if (!bmp) continue; // safety

                Image img;
                SoftwareBitmapSource src;
                try {
                    co_await src.SetBitmapAsync(bmp);
                }
                catch (hresult_error const& ex) {
                    Log(L"SetBitmapAsync failed: " + ex.message());
                    continue;
                }

                img.Source(src);
                img.Width(p.w);
                img.Height(p.h);
                img.Stretch(Stretch::Uniform);
                img.HorizontalAlignment(HorizontalAlignment::Left);
                img.VerticalAlignment(VerticalAlignment::Top);
                img.Margin({ p.x, p.y, 0, 0 });
                grid.Children().Append(img);
            }
            else
            {
                // Placeholder
                Border b;
                b.Background(SolidColorBrush(Microsoft::UI::Colors::LightGray()));
                b.BorderBrush(SolidColorBrush(Microsoft::UI::Colors::Gray()));
                b.BorderThickness({ 1,1,1,1 });
                b.Width(p.w);
                b.Height(p.h);
                b.HorizontalAlignment(HorizontalAlignment::Left);
                b.VerticalAlignment(VerticalAlignment::Top);
                b.Margin({ p.x, p.y, 0, 0 });

                if (p.rotated)
                {
                    TextBlock tb;
                    tb.Text(L"90\u00B0");
                    tb.HorizontalAlignment(HorizontalAlignment::Center);
                    tb.VerticalAlignment(VerticalAlignment::Center);
                    b.Child(tb);
                }

                grid.Children().Append(b);
            }
        }

        Log(L"Grid populated with " + to_hstring(static_cast<int>(m_currentPlacements.size())) + L" images.");
    }

    // ──────────────────────────────────────────────────────────────
    // Crop / Apply / Rotate bitmap helpers
    // ──────────────────────────────────────────────────────────────

    winrt::fire_and_forget MainWindow::BtnApplyCrop_Click(IInspectable const&, RoutedEventArgs const&)
    {
        Log(L"BtnApplyCrop Clicked.");
        if (!m_originalBitmap) co_return;

        auto stamp = co_await CaptureCropAsBitmap();
        if (!stamp) { Log(L"Failed to capture stamp."); co_return; }

        m_croppedStamp = stamp;
        m_croppedStampRotated = co_await RotateBitmap90(stamp);

        Log(L"Stamp captured (" + to_hstring(stamp.PixelWidth()) + L"x" + to_hstring(stamp.PixelHeight()) + L")");
        co_await RegeneratePreviewGrid();
    }

    winrt::Windows::Foundation::IAsyncOperation<SoftwareBitmap> MainWindow::CaptureCropAsBitmap()
    {
        if (!m_originalBitmap) co_return nullptr;

        double sourceW = static_cast<double>(m_originalBitmap.PixelWidth());
        double sourceH = static_cast<double>(m_originalBitmap.PixelHeight());

        auto scroller = CropScrollViewer();
        if (!scroller) co_return nullptr;

        float zoom = scroller.ZoomFactor();
        double hOffset  = scroller.HorizontalOffset();
        double vOffset  = scroller.VerticalOffset();
        double viewW    = scroller.ViewportWidth();
        double viewH    = scroller.ViewportHeight();

        double cropX = hOffset / zoom;
        double cropY = vOffset / zoom;
        double cropW = viewW  / zoom;
        double cropH = viewH  / zoom;

        // Clamp to source bounds
        cropX = std::max(0.0, cropX);
        cropY = std::max(0.0, cropY);
        if (cropX + cropW > sourceW) cropW = sourceW - cropX;
        if (cropY + cropH > sourceH) cropH = sourceH - cropY;
        if (cropW <= 0 || cropH <= 0) co_return nullptr;

        BitmapBounds bounds;
        bounds.X = static_cast<uint32_t>(cropX);
        bounds.Y = static_cast<uint32_t>(cropY);
        bounds.Width  = static_cast<uint32_t>(std::max(1.0, cropW));
        bounds.Height = static_cast<uint32_t>(std::max(1.0, cropH));

        Log(L"Crop: " + to_hstring(bounds.X) + L"," + to_hstring(bounds.Y) + L" "
            + to_hstring(bounds.Width) + L"x" + to_hstring(bounds.Height));

        try
        {
            InMemoryRandomAccessStream tmpStream;
            auto enc = co_await BitmapEncoder::CreateAsync(BitmapEncoder::PngEncoderId(), tmpStream);
            enc.SetSoftwareBitmap(m_originalBitmap);
            enc.BitmapTransform().Bounds(bounds);
            co_await enc.FlushAsync();

            tmpStream.Seek(0);
            auto dec = co_await BitmapDecoder::CreateAsync(tmpStream);
            auto croppedBmp = co_await dec.GetSoftwareBitmapAsync();

            if (croppedBmp.BitmapPixelFormat() != BitmapPixelFormat::Bgra8 ||
                croppedBmp.BitmapAlphaMode()   != BitmapAlphaMode::Premultiplied)
                croppedBmp = SoftwareBitmap::Convert(croppedBmp, BitmapPixelFormat::Bgra8, BitmapAlphaMode::Premultiplied);

            co_return croppedBmp;
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

    // ──────────────────────────────────────────────────────────────
    // Rotate source image 90° (replaces bitmap in-place)
    // ──────────────────────────────────────────────────────────────

    winrt::fire_and_forget MainWindow::BtnRotate_Click(IInspectable const&, RoutedEventArgs const&)
    {
        Log(L"Rotate Source Clicked.");
        if (!m_originalBitmap) co_return;

        auto strong = get_strong();          // prevent premature destruction

        try
        {
            auto rotated = co_await RotateBitmap90(m_originalBitmap);
            if (!rotated) { Log(L"Rotation failed."); co_return; }

            m_originalBitmap = rotated;

            SoftwareBitmapSource src;
            co_await src.SetBitmapAsync(m_originalBitmap);
            if (SourceImageControl()) SourceImageControl().Source(src);

            // Reset zoom to fit the new orientation
            ZoomToFit();

            // Invalidate stamps
            m_croppedStamp = nullptr;
            m_croppedStampRotated = nullptr;
            co_await RegeneratePreviewGrid();
            Log(L"Rotation complete.");
        }
        catch (hresult_error const& ex) {
            Log(L"Rotation exception: " + ex.message());
        }
    }

    // ──────────────────────────────────────────────────────────────
    // Zoom-to-fit: fits the loaded image inside the crop viewport
    // ──────────────────────────────────────────────────────────────

    void MainWindow::ZoomToFit()
    {
        try
        {
            auto scroller = CropScrollViewer();
            auto imgCtrl  = SourceImageControl();
            if (!scroller || !imgCtrl || !m_originalBitmap) return;

            double imageW = static_cast<double>(m_originalBitmap.PixelWidth());
            double imageH = static_cast<double>(m_originalBitmap.PixelHeight());
            if (imageW <= 0 || imageH <= 0) return;

            double viewportW = scroller.ViewportWidth();
            double viewportH = scroller.ViewportHeight();

            // If viewport isn't laid out yet, use container size
            if (viewportW <= 0 || viewportH <= 0)
            {
                auto container = CropContainer();
                if (container) {
                    viewportW = container.ActualWidth();
                    viewportH = container.ActualHeight();
                }
                if (viewportW <= 0 || viewportH <= 0) return;
            }

            float fitZoom = static_cast<float>(std::min(viewportW / imageW, viewportH / imageH));
            fitZoom = std::max(fitZoom, 0.1f);
            fitZoom = std::min(fitZoom, 10.0f);

            Log(L"ZoomToFit: img=" + to_hstring(imageW) + L"x" + to_hstring(imageH)
                + L" viewport=" + to_hstring(viewportW) + L"x" + to_hstring(viewportH)
                + L" zoom=" + to_hstring(fitZoom));

            m_zoomingFromMouse = true;
            scroller.ChangeView(0.0, 0.0, fitZoom);
            if (ZoomSlider()) ZoomSlider().Value(fitZoom);
            m_zoomingFromMouse = false;
        }
        catch (hresult_error const& ex) {
            m_zoomingFromMouse = false;
            Log(L"ZoomToFit failed: " + ex.message());
        }
    }

    // ──────────────────────────────────────────────────────────────
    // Save sheet at full 300 DPI resolution
    // ──────────────────────────────────────────────────────────────

    winrt::fire_and_forget MainWindow::BtnSaveSheet_Click(IInspectable const&, RoutedEventArgs const&)
    {
        auto grid = PreviewGrid();
        if (!grid) co_return;
        Log(L"Save Button Clicked.");

        auto strong = get_strong();

        FileSavePicker picker;
        auto initWnd{ picker.as<::IInitializeWithWindow>() };
        HWND hwnd;
        auto windowNative{ this->try_as<::IWindowNative>() };
        windowNative->get_WindowHandle(&hwnd);
        initWnd->Initialize(hwnd);
        picker.FileTypeChoices().Insert(L"JPEG", single_threaded_vector<hstring>({ L".jpg" }));
        picker.SuggestedFileName(L"PassportSheet");

        StorageFile file = co_await picker.PickSaveFileAsync();
        if (!file) co_return;

        Log(L"Saving to: " + file.Path());

        try
        {
            // RenderTargetBitmap has a max size limit (~16384 on most GPUs).
            // The PreviewGrid is already sized in 300-DPI pixels, so render
            // it at its actual layout size which the Viewbox has scaled down.
            // We request the grid's logical pixel dimensions.
            int targetW = static_cast<int>(grid.Width());
            int targetH = static_cast<int>(grid.Height());

            // Cap to safe maximum for RenderTargetBitmap
            const int kMaxRenderDim = 16384;
            if (targetW > kMaxRenderDim || targetH > kMaxRenderDim)
            {
                double scale = std::min(
                    static_cast<double>(kMaxRenderDim) / targetW,
                    static_cast<double>(kMaxRenderDim) / targetH);
                targetW = static_cast<int>(targetW * scale);
                targetH = static_cast<int>(targetH * scale);
                Log(L"Capped render size to " + to_hstring(targetW) + L"x" + to_hstring(targetH));
            }

            RenderTargetBitmap rtb;
            co_await rtb.RenderAsync(grid, targetW, targetH);

            SoftwareBitmap sb(BitmapPixelFormat::Bgra8, rtb.PixelWidth(), rtb.PixelHeight(), BitmapAlphaMode::Premultiplied);
            auto buffer = co_await rtb.GetPixelsAsync();
            sb.CopyFromBuffer(buffer);

            Log(L"Rendered: " + to_hstring(sb.PixelWidth()) + L"x" + to_hstring(sb.PixelHeight()));

            auto stream = co_await file.OpenAsync(FileAccessMode::ReadWrite);
            auto enc = co_await BitmapEncoder::CreateAsync(BitmapEncoder::JpegEncoderId(), stream);
            enc.SetSoftwareBitmap(sb);
            enc.BitmapTransform().InterpolationMode(BitmapInterpolationMode::Fant);
            co_await enc.FlushAsync();

            Log(L"Save Complete.");
        }
        catch (hresult_error const& ex) {
            Log(L"Save failed: " + ex.message());
        }
    }

    // ──────────────────────────────────────────────────────────────
    // Image loading (file picker + drag-and-drop)
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
        picker.FileTypeFilter().ReplaceAll({ L".jpg", L".png", L".jpeg", L".bmp", L".tiff" });
        StorageFile file = co_await picker.PickSingleFileAsync();
        if (file) co_await LoadImageFromFile(file);
    }

    winrt::Windows::Foundation::IAsyncAction MainWindow::LoadImageFromFile(StorageFile file)
    {
        auto strong = get_strong();
        Log(L"Loading Image: " + file.Path());

        try
        {
            auto stream = co_await file.OpenAsync(FileAccessMode::Read);
            auto decoder = co_await BitmapDecoder::CreateAsync(stream);
            m_originalBitmap = co_await decoder.GetSoftwareBitmapAsync();

            if (m_originalBitmap.BitmapPixelFormat() != BitmapPixelFormat::Bgra8 ||
                m_originalBitmap.BitmapAlphaMode()   != BitmapAlphaMode::Premultiplied)
            {
                m_originalBitmap = SoftwareBitmap::Convert(m_originalBitmap, BitmapPixelFormat::Bgra8, BitmapAlphaMode::Premultiplied);
            }

            SoftwareBitmapSource src;
            co_await src.SetBitmapAsync(m_originalBitmap);
            if (SourceImageControl()) SourceImageControl().Source(src);

            // Clear old stamps
            m_croppedStamp = nullptr;
            m_croppedStampRotated = nullptr;

            co_await RegeneratePreviewGrid();

            // Delay zoom-to-fit slightly to allow layout to complete
            auto dispatcher = this->DispatcherQueue();
            if (dispatcher)
            {
                auto weak = get_weak();
                dispatcher.TryEnqueue([weak]()
                {
                    if (auto self = weak.get())
                    {
                        self->ZoomToFit();
                    }
                });
            }
            Log(L"Image loaded. Size: " + to_hstring(m_originalBitmap.PixelWidth()) + L"x" + to_hstring(m_originalBitmap.PixelHeight()));
        }
        catch (hresult_error const& ex) {
            Log(L"LoadImageFromFile failed: " + ex.message());
        }
    }

    // ──────────────────────────────────────────────────────────────
    // Pointer input for crop panning + mouse-wheel zoom
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
        if (!m_isDragging || !CropScrollViewer()) return;

        auto pt = e.GetCurrentPoint(CropScrollViewer()).Position();
        double dx = pt.X - m_lastPoint.X;
        double dy = pt.Y - m_lastPoint.Y;

        CropScrollViewer().ChangeView(
            CropScrollViewer().HorizontalOffset() - dx,
            CropScrollViewer().VerticalOffset()  - dy,
            nullptr);

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
        double step = oldZoom * 0.1; // 10% step for smooth feel
        double newZoom = (delta > 0) ? oldZoom + step : oldZoom - step;
        newZoom = std::clamp(newZoom, 0.1, 10.0);

        if (newZoom == oldZoom) { e.Handled(true); return; }

        auto pt = e.GetCurrentPoint(scroller).Position();
        double hOff = scroller.HorizontalOffset();
        double vOff = scroller.VerticalOffset();

        // Zoom toward cursor position
        double newH = ((hOff + pt.X) / oldZoom) * newZoom - pt.X;
        double newV = ((vOff + pt.Y) / oldZoom) * newZoom - pt.Y;

        m_zoomingFromMouse = true;
        if (ZoomSlider()) ZoomSlider().Value(newZoom);
        scroller.ChangeView(newH, newV, static_cast<float>(newZoom));
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
