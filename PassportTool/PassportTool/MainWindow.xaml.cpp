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

        IncrementNumberRounder rounder;
        rounder.Increment(0.01);
        rounder.RoundingAlgorithm(RoundingAlgorithm::RoundHalfUp);
        formatter.NumberRounder(rounder);

        this->NbGap().NumberFormatter(formatter);
        this->NbSheetW().NumberFormatter(formatter);
        this->NbSheetH().NumberFormatter(formatter);

        this->ZoomSlider().ValueChanged([this](auto&&, auto&& args)
            {
                if (this->m_zoomingFromMouse) return;
                if (this->CropScrollViewer())
                {
                    this->CropScrollViewer().ChangeView(nullptr, nullptr, static_cast<float>(args.NewValue()));
                }
            });

        this->RegeneratePreviewGrid();
        this->RefitCropContainer();
    }

    int32_t MainWindow::MyProperty() { return m_myProperty; }
    void MainWindow::MyProperty(int32_t value) { m_myProperty = value; }

    double MainWindow::GetPixelsPerUnit()
    {
        double dpi = 300.0;
        if (this->RadioCm().IsChecked())
        {
            return dpi / 2.54;
        }
        return dpi;
    }

    void MainWindow::OnUnitChanged(IInspectable const&, RoutedEventArgs const&)
    {
        if (!this->NbSheetW() || !this->NbSheetH() || !this->NbGap()) return;

        auto radioChecked = this->RadioCm().IsChecked();
        bool toCm = (radioChecked != nullptr && radioChecked.Value());
        double factor = toCm ? 2.54 : (1.0 / 2.54);

        this->NbSheetW().Value(this->NbSheetW().Value() * factor);
        this->NbSheetH().Value(this->NbSheetH().Value() * factor);
        this->NbGap().Value(this->NbGap().Value() * factor);

        this->UpdateSheetSize();
        this->RegeneratePreviewGrid();
        this->RefitCropContainer();
    }

    void MainWindow::BtnRotate_Click(IInspectable const&, RoutedEventArgs const&)
    {
        double currentAngle = this->ImageRotation().Angle();
        this->ImageRotation().Angle(fmod(currentAngle + 90.0, 360.0));
    }

    void MainWindow::OnSettingsChanged(NumberBox const&, NumberBoxValueChangedEventArgs const&)
    {
        this->RegeneratePreviewGrid();
        this->RefitCropContainer();
    }

    void MainWindow::OnSheetDimChanged(NumberBox const&, NumberBoxValueChangedEventArgs const&)
    {
        this->UpdateSheetSize();
        this->RefitCropContainer();
    }

    void MainWindow::OnCropSizeChanged(IInspectable const&, SizeChangedEventArgs const&)
    {
        this->RefitCropContainer();
    }

    void MainWindow::RefitCropContainer()
    {
        if (!this->CropContainer() || !this->CropParentContainer()) return;
        double availableW = this->CropParentContainer().ActualWidth();
        double availableH = this->CropParentContainer().ActualHeight();
        if (availableW < 10 || availableH < 10) return;

        double targetRatio = (this->NbSheetW().Value() / this->NbCols().Value()) / (this->NbSheetH().Value() / this->NbRows().Value());
        double candidateW = availableH * targetRatio;
        double candidateH = availableH;

        if (candidateW > availableW)
        {
            candidateW = availableW;
            candidateH = availableW / targetRatio;
        }

        this->CropContainer().Width(candidateW);
        this->CropContainer().Height(candidateH);
    }

    void MainWindow::OnImagePointerWheelChanged(IInspectable const&, PointerRoutedEventArgs const& e)
    {
        int delta = e.GetCurrentPoint(this->CropScrollViewer()).Properties().MouseWheelDelta();
        double oldZoom = this->ZoomSlider().Value();
        double newZoom = (delta > 0) ? oldZoom + 0.1 : oldZoom - 0.1;

        if (newZoom < this->ZoomSlider().Minimum()) newZoom = this->ZoomSlider().Minimum();
        if (newZoom > this->ZoomSlider().Maximum()) newZoom = this->ZoomSlider().Maximum();
        if (newZoom == oldZoom) return;

        auto pt = e.GetCurrentPoint(this->CropScrollViewer()).Position();
        double hOff = this->CropScrollViewer().HorizontalOffset();
        double vOff = this->CropScrollViewer().VerticalOffset();

        double newH = ((hOff + pt.X) / oldZoom) * newZoom - pt.X;
        double newV = ((vOff + pt.Y) / oldZoom) * newZoom - pt.Y;

        this->m_zoomingFromMouse = true;
        this->ZoomSlider().Value(newZoom);
        this->CropScrollViewer().ChangeView(newH, newV, static_cast<float>(newZoom));
        this->m_zoomingFromMouse = false;
        e.Handled(true);
    }

    void MainWindow::OnImagePointerPressed(IInspectable const& sender, PointerRoutedEventArgs const& e)
    {
        auto img = sender.as<Image>();
        img.CapturePointer(e.Pointer());
        this->m_isDragging = true;
        this->m_lastPoint = e.GetCurrentPoint(this->CropScrollViewer()).Position();
    }

    void MainWindow::OnImagePointerMoved(IInspectable const&, PointerRoutedEventArgs const& e)
    {
        if (this->m_isDragging)
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
        sender.as<Image>().ReleasePointerCapture(e.Pointer());
    }

    void MainWindow::RegeneratePreviewGrid()
    {
        if (!this->PreviewGrid()) return;
        this->PreviewGrid().Children().Clear();
        this->PreviewGrid().RowDefinitions().Clear();
        this->PreviewGrid().ColumnDefinitions().Clear();

        int rows = static_cast<int>(this->NbRows().Value());
        int cols = static_cast<int>(this->NbCols().Value());
        double gapPx = this->NbGap().Value() * this->GetPixelsPerUnit();

        for (int i = 0; i < rows; ++i) { RowDefinition rd; rd.Height({ 1, GridUnitType::Star }); this->PreviewGrid().RowDefinitions().Append(rd); }
        for (int j = 0; j < cols; ++j) { ColumnDefinition cd; cd.Width({ 1, GridUnitType::Star }); this->PreviewGrid().ColumnDefinitions().Append(cd); }

        for (int r = 0; r < rows; ++r)
        {
            for (int c = 0; c < cols; ++c)
            {
                if (this->m_croppedPassportPhoto)
                {
                    Image img; img.Stretch(Stretch::UniformToFill); img.Margin({ gapPx / 2, gapPx / 2, gapPx / 2, gapPx / 2 });
                    SoftwareBitmapSource src; src.SetBitmapAsync(this->m_croppedPassportPhoto); img.Source(src);
                    Grid::SetRow(img, r); Grid::SetColumn(img, c); this->PreviewGrid().Children().Append(img);
                }
                else
                {
                    Border b; b.Background(SolidColorBrush(Microsoft::UI::Colors::LightGray())); b.Margin({ gapPx / 2, gapPx / 2, gapPx / 2, gapPx / 2 });
                    Grid::SetRow(b, r); Grid::SetColumn(b, c); this->PreviewGrid().Children().Append(b);
                }
            }
        }
    }

    void MainWindow::UpdateSheetSize()
    {
        if (this->PreviewGrid()) {
            this->PreviewGrid().Width(this->NbSheetW().Value() * this->GetPixelsPerUnit());
            this->PreviewGrid().Height(this->NbSheetH().Value() * this->GetPixelsPerUnit());
        }
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
        auto stream = co_await file.OpenAsync(FileAccessMode::Read);
        BitmapImage bmp;
        bmp.SetSource(stream);
        this->SourceImageControl().Source(bmp);
        this->ZoomSlider().Value(1.0);
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

    winrt::fire_and_forget MainWindow::BtnApplyCrop_Click(IInspectable const&, Microsoft::UI::Xaml::RoutedEventArgs const&)
    {
        if (this->SourceImageControl().Source()) {
            this->m_croppedPassportPhoto = co_await this->CaptureElementAsync(this->CropScrollViewer());
            this->RegeneratePreviewGrid();
        }
    }

    winrt::fire_and_forget MainWindow::BtnSaveSheet_Click(IInspectable const&, Microsoft::UI::Xaml::RoutedEventArgs const&)
    {
        if (!this->PreviewGrid()) co_return;
        auto bmp = co_await this->CaptureElementAsync(this->PreviewGrid());
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
            enc.SetSoftwareBitmap(bmp);
            co_await enc.FlushAsync();
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
}