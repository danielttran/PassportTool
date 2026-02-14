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
using namespace Windows::Storage;
using namespace Windows::Storage::Pickers;
using namespace Windows::Graphics::Imaging;

namespace winrt::PassportTool::implementation
{
    MainWindow::MainWindow()
    {
        InitializeComponent();
        RegeneratePreviewGrid();
    }

    int32_t MainWindow::MyProperty()
    {
        return m_myProperty;
    }

    void MainWindow::MyProperty(int32_t value)
    {
        m_myProperty = value;
    }

    void MainWindow::OnSettingsChanged(NumberBox const&, NumberBoxValueChangedEventArgs const&)
    {
        RegeneratePreviewGrid();
    }

    Windows::Foundation::IAsyncAction MainWindow::BtnPickImage_Click(IInspectable const&, RoutedEventArgs const&)
    {
        // 1. Create the picker
        FileOpenPicker picker;

        // 2. Initialize with Window Handle (Required for WinUI 3 Desktop)
        // We use the specific interface IInitializeWithWindow
        auto initializeWithWindow{ picker.as<::IInitializeWithWindow>() };

        // Retrieve the Window Handle (HWND) using IWindowNative
        auto windowNative{ this->try_as<::IWindowNative>() };
        HWND hWnd{ 0 };
        if (windowNative)
        {
            check_hresult(windowNative->get_WindowHandle(&hWnd));
        }
        initializeWithWindow->Initialize(hWnd);

        picker.ViewMode(PickerViewMode::Thumbnail);
        picker.SuggestedStartLocation(PickerLocationId::PicturesLibrary);
        picker.FileTypeFilter().Append(L".jpg");
        picker.FileTypeFilter().Append(L".png");
        picker.FileTypeFilter().Append(L".jpeg");

        // 3. Pick the file (Async)
        StorageFile file = co_await picker.PickSingleFileAsync();
        if (!file) co_return;

        // 4. Open stream
        auto stream = co_await file.OpenAsync(FileAccessMode::Read);

        // 5. Create Decoder
        BitmapDecoder decoder = co_await BitmapDecoder::CreateAsync(stream);

        // 6. Convert to SoftwareBitmap (BGRA8 Premultiplied)
        auto softwareBitmap = co_await decoder.GetSoftwareBitmapAsync();
        m_cachedSoftwareBitmap = SoftwareBitmap::Convert(softwareBitmap, BitmapPixelFormat::Bgra8, BitmapAlphaMode::Premultiplied);

        // 7. Show in the "Source" preview
        SoftwareBitmapSource source;
        co_await source.SetBitmapAsync(m_cachedSoftwareBitmap);
        SourceImagePreview().Source(source);

        // 8. Update the grid
        RegeneratePreviewGrid();
    }

    void MainWindow::RegeneratePreviewGrid()
    {
        if (!PreviewGrid()) return;

        // Clear old grid
        PreviewGrid().Children().Clear();
        PreviewGrid().RowDefinitions().Clear();
        PreviewGrid().ColumnDefinitions().Clear();

        // Get Inputs
        int rows = static_cast<int>(NbRows().Value());
        int cols = static_cast<int>(NbCols().Value());
        if (rows <= 0 || cols <= 0) return;

        // Create Grid Definitions
        for (int i = 0; i < rows; ++i)
        {
            RowDefinition rd;
            rd.Height(GridLength{ 1, GridUnitType::Star });
            PreviewGrid().RowDefinitions().Append(rd);
        }

        for (int j = 0; j < cols; ++j)
        {
            ColumnDefinition cd;
            cd.Width(GridLength{ 1, GridUnitType::Star });
            PreviewGrid().ColumnDefinitions().Append(cd);
        }

        // Populate Cells
        for (int row = 0; row < rows; ++row)
        {
            for (int col = 0; col < cols; ++col)
            {
                Image cellImage;
                cellImage.Stretch(Stretch::UniformToFill);
                cellImage.Margin({ 4, 4, 4, 4 });
                cellImage.HorizontalAlignment(HorizontalAlignment::Center);
                cellImage.VerticalAlignment(VerticalAlignment::Center);

                if (m_cachedSoftwareBitmap)
                {
                    SoftwareBitmapSource cellSource;
                    cellSource.SetBitmapAsync(m_cachedSoftwareBitmap);
                    cellImage.Source(cellSource);
                }
                else
                {
                    Border placeholder;
                    placeholder.Background(SolidColorBrush(Microsoft::UI::Colors::LightGray()));
                    Grid::SetRow(placeholder, row);
                    Grid::SetColumn(placeholder, col);
                    PreviewGrid().Children().Append(placeholder);
                    continue;
                }

                Grid::SetRow(cellImage, row);
                Grid::SetColumn(cellImage, col);
                PreviewGrid().Children().Append(cellImage);
            }
        }
    }
}