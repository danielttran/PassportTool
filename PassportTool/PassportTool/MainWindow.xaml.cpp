#include "pch.h"
#include "MainWindow.xaml.h"
#if __has_include("MainWindow.g.cpp")
#include "MainWindow.g.cpp"
#endif

using namespace winrt;
using namespace Microsoft::UI::Xaml;
using namespace Microsoft::UI::Xaml::Controls;
using namespace Microsoft::UI::Xaml::Media;
using namespace Microsoft::UI::Xaml::Media::Imaging;
using namespace Microsoft::UI::Xaml::Input;
using namespace Windows::Storage;
using namespace Windows::Storage::Pickers;
using namespace Windows::Graphics::Imaging;
using namespace Windows::Storage::Streams;
using namespace Windows::ApplicationModel::DataTransfer;
using namespace Windows::Globalization::NumberFormatting;

namespace winrt::PassportTool::implementation
{
    MainWindow::MainWindow()
    {
        InitializeComponent();

        DecimalFormatter formatter;
        formatter.IntegerDigits(1);
        formatter.FractionDigits(2);
        NbGap().NumberFormatter(formatter);

        ZoomSlider().ValueChanged([this](auto&&, auto&& args)
            {
                if (m_zoomingFromMouse) return;
                if (CropScrollViewer())
                {
                    CropScrollViewer().ChangeView(nullptr, nullptr, static_cast<float>(args.NewValue()));
                }
            });

        RegeneratePreviewGrid();
        RefitCropContainer();
    }

    int32_t MainWindow::MyProperty() { return m_myProperty; }
    void MainWindow::MyProperty(int32_t value) { m_myProperty = value; }

    void MainWindow::OnSettingsChanged(NumberBox const&, NumberBoxValueChangedEventArgs const&)
    {
        RegeneratePreviewGrid();
        RefitCropContainer();
    }

    void MainWindow::OnSheetDimChanged(NumberBox const&, NumberBoxValueChangedEventArgs const&)
    {
        UpdateSheetSize();
        RefitCropContainer();
    }

    void MainWindow::OnCropSizeChanged(IInspectable const&, SizeChangedEventArgs const&)
    {
        RefitCropContainer();
    }

    void MainWindow::RefitCropContainer()
    {
        if (!CropContainer() || !CropParentContainer()) return;
        double availableW = CropParentContainer().ActualWidth();
        double availableH = CropParentContainer().ActualHeight();
        if (availableW < 10 || availableH < 10) return;

        double sheetW = NbSheetW().Value();
        double sheetH = NbSheetH().Value();
        double rows = NbRows().Value();
        double cols = NbCols().Value();
        if (rows == 0 || cols == 0) return;

        double targetRatio = (sheetW / cols) / (sheetH / rows);
        double candidateW = availableH * targetRatio;
        double candidateH = availableH;

        if (candidateW > availableW)
        {
            candidateW = availableW;
            candidateH = availableW / targetRatio;
        }

        CropContainer().Width(candidateW);
        CropContainer().Height(candidateH);
    }

    void MainWindow::OnImagePointerWheelChanged(IInspectable const&, PointerRoutedEventArgs const& e)
    {
        int delta = e.GetCurrentPoint(CropScrollViewer()).Properties().MouseWheelDelta();
        double oldZoom = ZoomSlider().Value();
        double newZoom = (delta > 0) ? oldZoom + 0.1 : oldZoom - 0.1;

        if (newZoom < ZoomSlider().Minimum()) newZoom = ZoomSlider().Minimum();
        if (newZoom > ZoomSlider().Maximum()) newZoom = ZoomSlider().Maximum();
        if (newZoom == oldZoom) return;

        auto pt = e.GetCurrentPoint(CropScrollViewer()).Position();
        double hOff = CropScrollViewer().HorizontalOffset();
        double vOff = CropScrollViewer().VerticalOffset();

        double newH = ((hOff + pt.X) / oldZoom) * newZoom - pt.X;
        double newV = ((vOff + pt.Y) / oldZoom) * newZoom - pt.Y;

        m_zoomingFromMouse = true;
        ZoomSlider().Value(newZoom);
        CropScrollViewer().ChangeView(newH, newV, static_cast<float>(newZoom));
        m_zoomingFromMouse = false;
        e.Handled(true);
    }

    void MainWindow::OnImagePointerPressed(IInspectable const& sender, PointerRoutedEventArgs const& e)
    {
        auto img = sender.as<Image>();
        img.CapturePointer(e.Pointer());
        m_isDragging = true;
        m_lastPoint = e.GetCurrentPoint(CropScrollViewer()).Position();
    }

    void MainWindow::OnImagePointerMoved(IInspectable const&, PointerRoutedEventArgs const& e)
    {
        if (m_isDragging)
        {
            auto pt = e.GetCurrentPoint(CropScrollViewer()).Position();
            CropScrollViewer().ChangeView(CropScrollViewer().HorizontalOffset() - (pt.X - m_lastPoint.X),
                CropScrollViewer().VerticalOffset() - (pt.Y - m_lastPoint.Y), nullptr);
            m_lastPoint = pt;
        }
    }

    void MainWindow::OnImagePointerReleased(IInspectable const& sender, PointerRoutedEventArgs const& e)
    {
        m_isDragging = false;
        sender.as<Image>().ReleasePointerCapture(e.Pointer());
    }

    void MainWindow::RegeneratePreviewGrid()
    {
        if (!PreviewGrid()) return;
        PreviewGrid().Children().Clear();
        PreviewGrid().RowDefinitions().Clear();
        PreviewGrid().ColumnDefinitions().Clear();

        int rows = static_cast<int>(NbRows().Value());
        int cols = static_cast<int>(NbCols().Value());
        double gapPx = NbGap().Value() * 300.0;

        for (int i = 0; i < rows; ++i) { RowDefinition rd; rd.Height({ 1, GridUnitType::Star }); PreviewGrid().RowDefinitions().Append(rd); }
        for (int j = 0; j < cols; ++j) { ColumnDefinition cd; cd.Width({ 1, GridUnitType::Star }); PreviewGrid().ColumnDefinitions().Append(cd); }

        for (int r = 0; r < rows; ++r)
        {
            for (int c = 0; c < cols; ++c)
            {
                if (m_croppedPassportPhoto)
                {
                    Image img; img.Stretch(Stretch::UniformToFill); img.Margin({ gapPx / 2, gapPx / 2, gapPx / 2, gapPx / 2 });
                    SoftwareBitmapSource src; src.SetBitmapAsync(m_croppedPassportPhoto); img.Source(src);
                    Grid::SetRow(img, r); Grid::SetColumn(img, c); PreviewGrid().Children().Append(img);
                }
                else
                {
                    Border b; b.Background(SolidColorBrush(Microsoft::UI::Colors::LightGray())); b.Margin({ gapPx / 2, gapPx / 2, gapPx / 2, gapPx / 2 });
                    Grid::SetRow(b, r); Grid::SetColumn(b, c); PreviewGrid().Children().Append(b);
                }
            }
        }
    }

    void MainWindow::UpdateSheetSize()
    {
        if (PreviewGrid()) { PreviewGrid().Width(NbSheetW().Value() * 300.0); PreviewGrid().Height(NbSheetH().Value() * 300.0); }
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
            if (items.Size() > 0 && items.GetAt(0).IsOfType(StorageItemTypes::File)) co_await LoadImageFromFile(items.GetAt(0).as<StorageFile>());
        }
    }

    winrt::Windows::Foundation::IAsyncAction MainWindow::LoadImageFromFile(StorageFile file)
    {
        auto stream = co_await file.OpenAsync(FileAccessMode::Read);
        BitmapImage bmp; bmp.SetSource(stream); SourceImageControl().Source(bmp); ZoomSlider().Value(1.0);
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
        if (file) co_await LoadImageFromFile(file);
    }

    winrt::fire_and_forget MainWindow::BtnApplyCrop_Click(IInspectable const&, Microsoft::UI::Xaml::RoutedEventArgs const&)
    {
        if (SourceImageControl().Source()) { m_croppedPassportPhoto = co_await CaptureElementAsync(CropScrollViewer()); RegeneratePreviewGrid(); }
    }

    winrt::fire_and_forget MainWindow::BtnSaveSheet_Click(IInspectable const&, Microsoft::UI::Xaml::RoutedEventArgs const&)
    {
        if (!PreviewGrid()) co_return;
        auto bmp = co_await CaptureElementAsync(PreviewGrid());
        FileSavePicker picker;
        auto initializeWithWindow{ picker.as<::IInitializeWithWindow>() };
        HWND hwnd;
        auto windowNative{ this->try_as<::IWindowNative>() };
        windowNative->get_WindowHandle(&hwnd);
        initializeWithWindow->Initialize(hwnd);

        picker.FileTypeChoices().Insert(L"JPEG", single_threaded_vector<hstring>({ L".jpg" }));
        StorageFile file = co_await picker.PickSaveFileAsync();
        if (file)
        {
            auto stream = co_await file.OpenAsync(FileAccessMode::ReadWrite);
            BitmapEncoder enc = co_await BitmapEncoder::CreateAsync(BitmapEncoder::JpegEncoderId(), stream);
            enc.SetSoftwareBitmap(bmp); co_await enc.FlushAsync();
        }
    }

    winrt::Windows::Foundation::IAsyncOperation<SoftwareBitmap> MainWindow::CaptureElementAsync(UIElement el)
    {
        RenderTargetBitmap r; co_await r.RenderAsync(el);
        SoftwareBitmap sb(BitmapPixelFormat::Bgra8, r.PixelWidth(), r.PixelHeight(), BitmapAlphaMode::Premultiplied);
        auto pixels = co_await r.GetPixelsAsync();
        sb.CopyFromBuffer(pixels); co_return sb;
    }
}