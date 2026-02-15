#include "pch.h"
#include "MainWindow.xaml.h"
#if __has_include("MainWindow.g.cpp")
#include "MainWindow.g.cpp"
#endif
#include <cmath>
#include <algorithm>

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
    MainWindow::MainWindow()
    {
        InitializeComponent();

        m_isLoaded = false;
        Log(L"MainWindow Constructor started.");

        // Configure Number Formatters
        DecimalFormatter formatter;
        formatter.IntegerDigits(1);
        formatter.FractionDigits(2);
        IncrementNumberRounder rounder;
        rounder.Increment(0.01);
        rounder.RoundingAlgorithm(RoundingAlgorithm::RoundHalfUp);
        formatter.NumberRounder(rounder);

        if (this->NbGap()) this->NbGap().NumberFormatter(formatter);
        if (this->NbSheetW()) this->NbSheetW().NumberFormatter(formatter);
        if (this->NbSheetH()) this->NbSheetH().NumberFormatter(formatter);
        if (this->NbImageW()) this->NbImageW().NumberFormatter(formatter);
        if (this->NbImageH()) this->NbImageH().NumberFormatter(formatter);

        if (this->ZoomSlider()) {
            this->ZoomSlider().ValueChanged([this](auto&&, auto&& args)
                {
                    if (this->m_zoomingFromMouse) return;
                    if (this->CropScrollViewer())
                    {
                        this->CropScrollViewer().ChangeView(nullptr, nullptr, static_cast<float>(args.NewValue()));
                    }
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
        this->UpdateSheetSize();
        this->RefitCropContainer();
        this->UpdateCellDimensionsDisplay();
        this->RegeneratePreviewGrid();
    }

    void MainWindow::UpdateCellDimensionsDisplay()
    {
        if (!m_isLoaded) return;

        auto imgWBox = this->NbImageW();
        auto imgHBox = this->NbImageH();
        auto txtCell = this->TxtCellDimensions();

        if (!imgWBox || !imgHBox || !txtCell) return;

        double imgW = imgWBox.Value();
        double imgH = imgHBox.Value();

        bool isCm = false;
        if (this->RadioCm()) {
            auto check = this->RadioCm().IsChecked();
            if (check) isCm = check.Value();
        }

        hstring unit = isCm ? L"cm" : L"in";
        // Use Unicode multiplication sign (U+00D7)
        hstring cellText = L"Cell: " + to_hstring(imgW) + L" x " + to_hstring(imgH) + L" " + unit;
        txtCell.Text(cellText);
    }

    double MainWindow::GetPixelsPerUnit()
    {
        double dpi = 300.0;
        if (this->RadioCm())
        {
            auto isChecked = this->RadioCm().IsChecked();
            if (isChecked && isChecked.Value()) return dpi / 2.54;
        }
        return dpi;
    }

    void MainWindow::OnUnitChanged(IInspectable const&, RoutedEventArgs const&)
    {
        if (!m_isLoaded) return;

        bool toCm = false;
        if (this->RadioCm()) {
            auto check = this->RadioCm().IsChecked();
            if (check) toCm = check.Value();
        }

        Log(L"Unit Changed. Converting to " + (toCm ? hstring(L"CM") : hstring(L"INCHES")));

        double factor = toCm ? 2.54 : (1.0 / 2.54);
        Log(L"Conversion factor: " + to_hstring(factor));

        auto convert = [&](NumberBox const& box) {
            if (box) {
                double oldVal = box.Value();
                box.Value(oldVal * factor);
                Log(L"  Converted: " + to_hstring(oldVal) + L" -> " + to_hstring(box.Value()));
            }
        };

        m_isLoaded = false;
        Log(L"Before conversion:");
        Log(L"  SheetW: " + to_hstring(this->NbSheetW() ? this->NbSheetW().Value() : 0));
        Log(L"  SheetH: " + to_hstring(this->NbSheetH() ? this->NbSheetH().Value() : 0));
        Log(L"  ImageW: " + to_hstring(this->NbImageW() ? this->NbImageW().Value() : 0));
        Log(L"  ImageH: " + to_hstring(this->NbImageH() ? this->NbImageH().Value() : 0));
        Log(L"  Gap: " + to_hstring(this->NbGap() ? this->NbGap().Value() : 0));

        convert(this->NbSheetW());
        convert(this->NbSheetH());
        convert(this->NbImageW());
        convert(this->NbImageH());
        convert(this->NbGap());

        Log(L"After conversion:");
        Log(L"  SheetW: " + to_hstring(this->NbSheetW() ? this->NbSheetW().Value() : 0));
        Log(L"  SheetH: " + to_hstring(this->NbSheetH() ? this->NbSheetH().Value() : 0));
        Log(L"  ImageW: " + to_hstring(this->NbImageW() ? this->NbImageW().Value() : 0));
        Log(L"  ImageH: " + to_hstring(this->NbImageH() ? this->NbImageH().Value() : 0));
        Log(L"  Gap: " + to_hstring(this->NbGap() ? this->NbGap().Value() : 0));

        m_isLoaded = true;

        this->UpdateSheetSize();
        this->RefitCropContainer();
        this->UpdateCellDimensionsDisplay();
        this->RegeneratePreviewGrid();
    }

    void MainWindow::OnSettingsChanged(NumberBox const& sender, NumberBoxValueChangedEventArgs const&)
    {
        if (!m_isLoaded) return;

        hstring header = L"Unknown";
        if (sender) header = sender.Header().as<hstring>();
        double val = sender ? sender.Value() : 0.0;
        Log(L"Setting Changed: " + header + L" = " + to_hstring(val));

        this->UpdateSheetSize();
        this->RefitCropContainer();
        this->UpdateCellDimensionsDisplay();
        this->RegeneratePreviewGrid();
    }

    void MainWindow::UpdateSheetSize()
    {
        if (!m_isLoaded) return;
        auto grid = this->PreviewGrid();
        auto sw = this->NbSheetW();
        auto sh = this->NbSheetH();

        if (!grid || !sw || !sh) return;

        double w = sw.Value() * this->GetPixelsPerUnit();
        double h = sh.Value() * this->GetPixelsPerUnit();

        if (w < 1) w = 1;
        if (h < 1) h = 1;

        Log(L"Sheet Updated: " + to_hstring(w) + L"x" + to_hstring(h) + L" px");
        grid.Width(w);
        grid.Height(h);
    }

    void MainWindow::RefitCropContainer()
    {
        auto container = this->CropContainer();
        auto parent = this->CropParentContainer();
        auto imgWBox = this->NbImageW();
        auto imgHBox = this->NbImageH();

        if (!container || !parent || !imgWBox || !imgHBox) return;

        double availableW = parent.ActualWidth();
        double availableH = parent.ActualHeight();

        if (availableW < 10 || availableH < 10) return;

        double imgW = imgWBox.Value();
        double imgH = imgHBox.Value();
        if (imgW <= 0 || imgH <= 0) return;

        double targetRatio = imgW / imgH;

        double candidateW = availableW;
        double candidateH = availableW / targetRatio;

        if (candidateH > availableH)
        {
            candidateH = availableH;
            candidateW = availableH * targetRatio;
        }

        container.Width(candidateW - 4);
        container.Height(candidateH - 4);
    }

    void MainWindow::OnCropSizeChanged(IInspectable const&, SizeChangedEventArgs const&)
    {
        this->RefitCropContainer();
    }

    winrt::Windows::Foundation::IAsyncAction MainWindow::RegeneratePreviewGrid()
    {
        if (!m_isLoaded) co_return;
        Log(L"Regenerating Preview Grid...");

        auto grid = this->PreviewGrid();
        auto swBox = this->NbSheetW();
        auto shBox = this->NbSheetH();
        auto imgWBox = this->NbImageW();
        auto imgHBox = this->NbImageH();
        auto gapBox = this->NbGap();
        auto txtInfo = this->TxtLayoutInfo();

        if (!grid || !swBox || !shBox || !imgWBox || !imgHBox || !gapBox) {
            Log(L"Missing controls for grid generation");
            Log(L"  grid=" + to_hstring(grid != nullptr) + L" swBox=" + to_hstring(swBox != nullptr) 
                + L" shBox=" + to_hstring(shBox != nullptr) + L" imgWBox=" + to_hstring(imgWBox != nullptr)
                + L" imgHBox=" + to_hstring(imgHBox != nullptr) + L" gapBox=" + to_hstring(gapBox != nullptr));
            co_return;
        }

        grid.Children().Clear();
        grid.RowDefinitions().Clear();
        grid.ColumnDefinitions().Clear();
        Log(L"Grid cleared");

        double sheetW = swBox.Value();
        double sheetH = shBox.Value();
        double imgW = imgWBox.Value();
        double imgH = imgHBox.Value();
        double gap = gapBox.Value();

        Log(L"Sheet: " + to_hstring(sheetW) + L"x" + to_hstring(sheetH) 
            + L" | Image: " + to_hstring(imgW) + L"x" + to_hstring(imgH) 
            + L" | Gap: " + to_hstring(gap));

        if (sheetW <= 0 || sheetH <= 0 || imgW <= 0 || imgH <= 0) {
            Log(L"Invalid dimensions");
            co_return;
        }

        // Formula: N * (Size + Gap) <= Sheet - Gap
        auto calc = [&](double dim, double size, double g) -> int {
            if (dim < g) return 0;
            return static_cast<int>(floor((dim - g) / (size + g)));
            };

        int colsA = calc(sheetW, imgW, gap);
        int rowsA = calc(sheetH, imgH, gap);
        int countA = colsA * rowsA;

        int colsB = calc(sheetW, imgH, gap);
        int rowsB = calc(sheetH, imgW, gap);
        int countB = colsB * rowsB;

        bool useRotation = (countB > countA);
        int bestCols = useRotation ? colsB : colsA;
        int bestRows = useRotation ? rowsB : rowsA;
        int total = useRotation ? countB : countA;

        Log(L"Layout Calc: Normal=" + to_hstring(countA) + L" Rotated=" + to_hstring(countB) + L" Chosen=" + (useRotation ? L"Rotated" : L"Normal"));
        Log(L"Best layout: " + to_hstring(bestRows) + L" rows x " + to_hstring(bestCols) + L" cols = " + to_hstring(total) + L" images");

        if (bestRows <= 0 || bestCols <= 0) {
            Log(L"No valid layout calculated");
            if (txtInfo) {
                txtInfo.Text(L"Images don't fit on sheet");
            }
            co_return;
        }

        if (txtInfo) {
            txtInfo.Text(to_hstring(total) + L" images (" + to_hstring(bestRows) + L" x " + to_hstring(bestCols) + L")" + (useRotation ? L" Rotated" : L""));
        }

        double ppu = this->GetPixelsPerUnit();
        double gapPx = gap * ppu;

        double realImgW_px = imgW * ppu;
        double realImgH_px = imgH * ppu;

        double cellW_px = (useRotation ? imgH : imgW) * ppu;
        double cellH_px = (useRotation ? imgW : imgH) * ppu;

        // Ensure image aspect ratio matches the settings
        double aspectW = imgW * ppu;
        double aspectH = imgH * ppu;

        // When rendering, set image width and height to match aspect ratio
        // This ensures the sheet shows 4 x 6 cm images, not squares

        Log(L"Pixels per unit: " + to_hstring(ppu) + L" | Cell: " + to_hstring(cellW_px) + L"x" + to_hstring(cellH_px));
        Log(L"Gap in pixels: " + to_hstring(gapPx));
        Log(L"Stamp info: " + to_hstring(m_croppedStamp != nullptr ? L"has stamp" : L"NO STAMP"));

        for (int i = 0; i < bestRows; ++i) {
            RowDefinition rd;
            rd.Height({ cellH_px, GridUnitType::Pixel });
            grid.RowDefinitions().Append(rd);
        }
        for (int j = 0; j < bestCols; ++j) {
            ColumnDefinition cd;
            cd.Width({ cellW_px, GridUnitType::Pixel });
            grid.ColumnDefinitions().Append(cd);
        }

        grid.RowSpacing(gapPx);
        grid.ColumnSpacing(gapPx);
        grid.Padding({ gapPx, gapPx, gapPx, gapPx });

        grid.HorizontalAlignment(HorizontalAlignment::Left);
        grid.VerticalAlignment(VerticalAlignment::Top);

        Log(L"Created grid structure: " + to_hstring(bestRows) + L" rows, " + to_hstring(bestCols) + L" cols");

        int cellCount = 0;
        for (int r = 0; r < bestRows; ++r)
        {
            for (int c = 0; c < bestCols; ++c)
            {
                Log(L"Creating cell [" + to_hstring(r) + L"," + to_hstring(c) + L"]");

                if (this->m_croppedStamp)
                {
                    Log(L"  Stamp exists, creating image");

                    Image img;
                    SoftwareBitmapSource src;

                    Log(L"  Created Image and SoftwareBitmapSource objects");

                    try 
                    {
                        Log(L"  Calling SetBitmapAsync...");
                        co_await src.SetBitmapAsync(this->m_croppedStamp);
                        Log(L"  SetBitmapAsync completed");
                    }
                    catch (hresult_error const& ex)
                    {
                        Log(L"  SetBitmapAsync failed: " + ex.message());
                        co_return;
                    }

                    img.Source(src);
                    Log(L"  Image source set");

                    // Always set width and height to match user aspect ratio
                    if (useRotation)
                    {
                        img.Width(realImgH_px); // swap for rotation
                        img.Height(realImgW_px);
                        img.Stretch(Stretch::Uniform);
                        img.RenderTransformOrigin({ 0.5, 0.5 });
                        RotateTransform rot;
                        rot.Angle(90);
                        img.RenderTransform(rot);
                        Log(L"  Applied 90° rotation, Stretch=Uniform");
                    }
                    else
                    {
                        img.Width(realImgW_px);
                        img.Height(realImgH_px);
                        img.Stretch(Stretch::Uniform);
                        Log(L"  No rotation, Stretch=Uniform");
                    }

                    Grid::SetRow(img, r);
                    Grid::SetColumn(img, c);
                    grid.Children().Append(img);
                    cellCount++;
                    Log(L"  Image added to grid");
                }
                else
                {
                    Log(L"  No stamp, creating placeholder border");

                    Border b;
                    b.Background(SolidColorBrush(Microsoft::UI::Colors::LightGray()));
                    b.BorderBrush(SolidColorBrush(Microsoft::UI::Colors::Gray()));
                    b.BorderThickness({ 1,1,1,1 });

                    if (useRotation) {
                        TextBlock tb; tb.Text(L"90°");
                        tb.HorizontalAlignment(HorizontalAlignment::Center); tb.VerticalAlignment(VerticalAlignment::Center);
                        b.Child(tb);
                    }

                    Grid::SetRow(b, r);
                    Grid::SetColumn(b, c);
                    grid.Children().Append(b);
                    cellCount++;
                    Log(L"  Placeholder added to grid");
                }
            }
        }

        Log(L"Grid population complete. Total cells added: " + to_hstring(cellCount));
        Log(L"Grid children count: " + to_hstring(grid.Children().Size()));
        Log(L"Grid row count: " + to_hstring(grid.RowDefinitions().Size()));
        Log(L"Grid column count: " + to_hstring(grid.ColumnDefinitions().Size()));
    }

    winrt::fire_and_forget MainWindow::BtnApplyCrop_Click(IInspectable const&, Microsoft::UI::Xaml::RoutedEventArgs const&)
    {
        Log(L"BtnApplyCrop Clicked.");
        if (this->m_originalBitmap) {
            auto stamp = co_await this->CaptureCropAsBitmap();
            if (stamp) {
                this->m_croppedStamp = stamp;
                Log(L"Stamp captured successfully.");
                co_await this->RegeneratePreviewGrid();
            }
            else {
                Log(L"Failed to capture stamp.");
            }
        }
    }

    winrt::Windows::Foundation::IAsyncOperation<SoftwareBitmap> MainWindow::CaptureCropAsBitmap()
    {
        if (!m_originalBitmap) co_return nullptr;

        double sourceW = static_cast<double>(m_originalBitmap.PixelWidth());
        double sourceH = static_cast<double>(m_originalBitmap.PixelHeight());

        auto scroller = this->CropScrollViewer();
        if (!scroller) co_return nullptr;

        float zoom = scroller.ZoomFactor();
        double hOffset = scroller.HorizontalOffset();
        double vOffset = scroller.VerticalOffset();
        double viewportW = scroller.ViewportWidth();
        double viewportH = scroller.ViewportHeight();

        double cropX = hOffset / zoom;
        double cropY = vOffset / zoom;
        double cropW = viewportW / zoom;
        double cropH = viewportH / zoom;

        Log(L"Crop Calc: Offset(" + to_hstring(hOffset) + L"," + to_hstring(vOffset) + L") Zoom=" + to_hstring(zoom));
        Log(L"Crop Rect Raw: " + to_hstring(cropX) + L"," + to_hstring(cropY) + L" " + to_hstring(cropW) + L"x" + to_hstring(cropH));

        if (cropX < 0) cropX = 0;
        if (cropY < 0) cropY = 0;
        if (cropX + cropW > sourceW) cropW = sourceW - cropX;
        if (cropY + cropH > sourceH) cropH = sourceH - cropY;

        if (cropW <= 0 || cropH <= 0) {
            Log(L"Invalid Crop dimensions.");
            co_return nullptr;
        }

        BitmapBounds extractBounds;
        extractBounds.X = static_cast<uint32_t>(cropX);
        extractBounds.Y = static_cast<uint32_t>(cropY);
        extractBounds.Width = static_cast<uint32_t>(cropW);
        extractBounds.Height = static_cast<uint32_t>(cropH);

        Log(L"Extracting region: X=" + to_hstring(extractBounds.X) + L" Y=" + to_hstring(extractBounds.Y) 
            + L" W=" + to_hstring(extractBounds.Width) + L" H=" + to_hstring(extractBounds.Height));

        try
        {
            InMemoryRandomAccessStream tempStream;
            auto encoder = co_await BitmapEncoder::CreateAsync(BitmapEncoder::PngEncoderId(), tempStream);
            encoder.SetSoftwareBitmap(m_originalBitmap);
            encoder.BitmapTransform().Bounds(extractBounds);
            co_await encoder.FlushAsync();

            tempStream.Seek(0);
            auto decoder = co_await BitmapDecoder::CreateAsync(tempStream);
            auto croppedBitmap = co_await decoder.GetSoftwareBitmapAsync();

            if (croppedBitmap.BitmapPixelFormat() != BitmapPixelFormat::Bgra8 || 
                croppedBitmap.BitmapAlphaMode() != BitmapAlphaMode::Premultiplied)
            {
                croppedBitmap = SoftwareBitmap::Convert(croppedBitmap, BitmapPixelFormat::Bgra8, BitmapAlphaMode::Premultiplied);
            }

            Log(L"Crop succeeded: " + to_hstring(croppedBitmap.PixelWidth()) + L"x" + to_hstring(croppedBitmap.PixelHeight()));
            co_return croppedBitmap;
        }
        catch (hresult_error const& ex)
        {
            Log(L"Crop Failed: " + ex.message());
            co_return nullptr;
        }
    }

    winrt::Windows::Foundation::IAsyncOperation<SoftwareBitmap> MainWindow::GenerateFullResolutionSheet()
    {
        // This function is deprecated - use RenderTargetBitmap on PreviewGrid instead
        // Kept for reference/future use
        if (!m_croppedStamp) {
            Log(L"GenerateFullResolutionSheet: No stamped image available.");
            co_return nullptr;
        }

        Log(L"GenerateFullResolutionSheet: This function is deprecated. Rendering is now done directly from PreviewGrid.");
        co_return nullptr;
    }

    winrt::fire_and_forget MainWindow::BtnSaveSheet_Click(IInspectable const&, Microsoft::UI::Xaml::RoutedEventArgs const&)
    {
        if (!this->PreviewGrid()) co_return;
        Log(L"Save Button Clicked.");

        FileSavePicker picker;
        auto initializeWithWindow{ picker.as<::IInitializeWithWindow>() };
        HWND hwnd;
        auto windowNative{ this->try_as<::IWindowNative>() };
        windowNative->get_WindowHandle(&hwnd);
        initializeWithWindow->Initialize(hwnd);
        picker.FileTypeChoices().Insert(L"JPEG", single_threaded_vector<hstring>({ L".jpg" }));
        picker.SuggestedFileName(L"PassportSheet");

        StorageFile file = co_await picker.PickSaveFileAsync();
        if (file)
        {
            Log(L"File Picked: " + file.Path());

            // Render the PreviewGrid directly to get the layout with all images
            try
            {
                auto previewGrid = this->PreviewGrid();
                if (!previewGrid)
                {
                    Log(L"PreviewGrid not found.");
                    co_return;
                }

                Log(L"Rendering PreviewGrid to bitmap...");
                RenderTargetBitmap rtb;
                co_await rtb.RenderAsync(previewGrid);

                Log(L"Creating SoftwareBitmap from rendered bitmap...");
                SoftwareBitmap sb(BitmapPixelFormat::Bgra8, rtb.PixelWidth(), rtb.PixelHeight(), BitmapAlphaMode::Premultiplied);
                auto buffer = co_await rtb.GetPixelsAsync();
                sb.CopyFromBuffer(buffer);

                Log(L"Rendered bitmap: " + to_hstring(sb.PixelWidth()) + L"x" + to_hstring(sb.PixelHeight()));

                auto stream = co_await file.OpenAsync(FileAccessMode::ReadWrite);
                BitmapEncoder enc = co_await BitmapEncoder::CreateAsync(BitmapEncoder::JpegEncoderId(), stream);
                enc.SetSoftwareBitmap(sb);
                co_await enc.FlushAsync();
                Log(L"Save Complete.");
            }
            catch (hresult_error const& ex)
            {
                Log(L"Save failed: " + ex.message());
            }
        }
    }

    winrt::Windows::Foundation::IAsyncOperation<SoftwareBitmap> MainWindow::CaptureElementAsync(UIElement el)
    {
        RenderTargetBitmap r;
        co_await r.RenderAsync(el);
        SoftwareBitmap sb(BitmapPixelFormat::Bgra8, r.PixelWidth(), r.PixelHeight(), BitmapAlphaMode::Premultiplied);
        auto buffer = co_await r.GetPixelsAsync();
        sb.CopyFromBuffer(buffer);
        co_return sb;
    }

    void MainWindow::BtnRotate_Click(IInspectable const&, RoutedEventArgs const&)
    {
        Log(L"Rotate Source Clicked.");
        auto rotateOp = [this]() -> winrt::Windows::Foundation::IAsyncAction {
            if (!this->m_originalBitmap) co_return;

            InMemoryRandomAccessStream stream;
            auto encoder = co_await BitmapEncoder::CreateAsync(BitmapEncoder::PngEncoderId(), stream);
            encoder.SetSoftwareBitmap(this->m_originalBitmap);

            encoder.BitmapTransform().Rotation(BitmapRotation::Clockwise90Degrees);

            try {
                co_await encoder.FlushAsync();
                stream.Seek(0);
                auto decoder = co_await BitmapDecoder::CreateAsync(stream);
                this->m_originalBitmap = co_await decoder.GetSoftwareBitmapAsync();

                if (this->m_originalBitmap.BitmapPixelFormat() != BitmapPixelFormat::Bgra8 ||
                    this->m_originalBitmap.BitmapAlphaMode() != BitmapAlphaMode::Premultiplied)
                {
                    this->m_originalBitmap = SoftwareBitmap::Convert(this->m_originalBitmap, BitmapPixelFormat::Bgra8, BitmapAlphaMode::Premultiplied);
                }

                SoftwareBitmapSource src;
                co_await src.SetBitmapAsync(this->m_originalBitmap);

                if (this->SourceImageControl()) this->SourceImageControl().Source(src);
                if (this->ZoomSlider()) this->ZoomSlider().Value(1.0);
                if (this->ImageRotation()) this->ImageRotation().Angle(0);

                this->m_croppedStamp = nullptr;
                co_await this->RegeneratePreviewGrid();
                Log(L"Rotation Complete.");
            }
            catch (...) { Log(L"Rotation Failed."); }
            };
        rotateOp();
    }

    void MainWindow::OnImagePointerWheelChanged(IInspectable const&, PointerRoutedEventArgs const& e)
    {
        if (!this->CropScrollViewer()) return;

        int delta = e.GetCurrentPoint(this->CropScrollViewer()).Properties().MouseWheelDelta();

        double oldZoom = 1.0;
        if (this->ZoomSlider()) oldZoom = this->ZoomSlider().Value();

        double newZoom = (delta > 0) ? oldZoom + 0.1 : oldZoom - 0.1;

        if (this->ZoomSlider()) {
            if (newZoom < this->ZoomSlider().Minimum()) newZoom = this->ZoomSlider().Minimum();
            if (newZoom > this->ZoomSlider().Maximum()) newZoom = this->ZoomSlider().Maximum();
        }

        if (newZoom == oldZoom) return;

        auto pt = e.GetCurrentPoint(this->CropScrollViewer()).Position();
        double hOff = this->CropScrollViewer().HorizontalOffset();
        double vOff = this->CropScrollViewer().VerticalOffset();

        double newH = ((hOff + pt.X) / oldZoom) * newZoom - pt.X;
        double newV = ((vOff + pt.Y) / oldZoom) * newZoom - pt.Y;

        this->m_zoomingFromMouse = true;
        if (this->ZoomSlider()) this->ZoomSlider().Value(newZoom);
        this->CropScrollViewer().ChangeView(newH, newV, static_cast<float>(newZoom));
        this->m_zoomingFromMouse = false;
        e.Handled(true);
    }

    void MainWindow::OnImagePointerPressed(IInspectable const& sender, PointerRoutedEventArgs const& e)
    {
        auto img = sender.as<Image>();
        if (!img) return;
        img.CapturePointer(e.Pointer());
        this->m_isDragging = true;
        if (this->CropScrollViewer())
            this->m_lastPoint = e.GetCurrentPoint(this->CropScrollViewer()).Position();
    }

    void MainWindow::OnImagePointerMoved(IInspectable const&, PointerRoutedEventArgs const& e)
    {
        if (this->m_isDragging && this->CropScrollViewer())
        {
            auto pt = e.GetCurrentPoint(this->CropScrollViewer()).Position();
            this->CropScrollViewer().ChangeView(this->CropScrollViewer().HorizontalOffset() - (pt.X - this->m_lastPoint.X),
                this->CropScrollViewer().VerticalOffset() - (pt.Y - this->m_lastPoint.Y), nullptr);
            this->m_lastPoint = pt;
        }
    }

    void MainWindow::OnImagePointerReleased(IInspectable const& sender, PointerRoutedEventArgs const& e)
    {
        this->m_isDragging = false;
        auto img = sender.try_as<Image>();
        if (img) img.ReleasePointerCapture(e.Pointer());
    }

    void MainWindow::OnDragOver(IInspectable const&, DragEventArgs const& e)
    {
        if (e.DataView().Contains(StandardDataFormats::StorageItems())) e.AcceptedOperation(DataPackageOperation::Copy);
    }

    winrt::fire_and_forget MainWindow::OnDrop(IInspectable const&, DragEventArgs const& e)
    {
        if (e.DataView().Contains(StandardDataFormats::StorageItems()))
        {
            auto items = co_await e.DataView().GetStorageItemsAsync();
            if (items.Size() > 0 && items.GetAt(0).IsOfType(StorageItemTypes::File)) {
                co_await this->LoadImageFromFile(items.GetAt(0).as<winrt::Windows::Storage::StorageFile>());
            }
        }
    }

    winrt::Windows::Foundation::IAsyncAction MainWindow::LoadImageFromFile(winrt::Windows::Storage::StorageFile file)
    {
        Log(L"Loading Image: " + file.Path());
        auto stream = co_await file.OpenAsync(FileAccessMode::Read);

        auto decoder = co_await BitmapDecoder::CreateAsync(stream);
        this->m_originalBitmap = co_await decoder.GetSoftwareBitmapAsync();

        if (m_originalBitmap.BitmapPixelFormat() != BitmapPixelFormat::Bgra8 || m_originalBitmap.BitmapAlphaMode() != BitmapAlphaMode::Premultiplied)
        {
            m_originalBitmap = SoftwareBitmap::Convert(m_originalBitmap, BitmapPixelFormat::Bgra8, BitmapAlphaMode::Premultiplied);
        }

        SoftwareBitmapSource src;
        co_await src.SetBitmapAsync(m_originalBitmap);
        if (this->SourceImageControl()) this->SourceImageControl().Source(src);

        if (this->ZoomSlider()) this->ZoomSlider().Value(1.0);
        if (this->ImageRotation()) this->ImageRotation().Angle(0);

        this->m_croppedStamp = nullptr;
        co_await this->RegeneratePreviewGrid();
    }

    winrt::fire_and_forget MainWindow::BtnPickImage_Click(IInspectable const&, Microsoft::UI::Xaml::RoutedEventArgs const&)
    {
        FileOpenPicker picker;
        auto initializeWithWindow{ picker.as<::IInitializeWithWindow>() };
        HWND hwnd;
        auto windowNative{ this->try_as<::IWindowNative>() };
        windowNative->get_WindowHandle(&hwnd);
        initializeWithWindow->Initialize(hwnd);
        picker.FileTypeFilter().ReplaceAll({ L".jpg", L".png", L".jpeg" });
        StorageFile file = co_await picker.PickSingleFileAsync();
        if (file) co_await this->LoadImageFromFile(file);
    }
}
