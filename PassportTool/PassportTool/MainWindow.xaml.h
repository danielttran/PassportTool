#pragma once

#include "MainWindow.g.h"
// Ensure these are visible in the header as well
#include <winrt/Windows.Graphics.Imaging.h>
#include <winrt/Microsoft.UI.Xaml.Media.Imaging.h>

namespace winrt::PassportTool::implementation
{
    struct MainWindow : MainWindowT<MainWindow>
    {
        MainWindow();

        int32_t MyProperty();
        void MyProperty(int32_t value);

        // Event Handlers
        Windows::Foundation::IAsyncAction BtnPickImage_Click(Windows::Foundation::IInspectable const& sender, Microsoft::UI::Xaml::RoutedEventArgs const& e);
        void OnSettingsChanged(Microsoft::UI::Xaml::Controls::NumberBox const& sender, Microsoft::UI::Xaml::Controls::NumberBoxValueChangedEventArgs const& args);

    private:
        // Helper to rebuild grid
        void RegeneratePreviewGrid();

        // Cached bitmap data
        Windows::Graphics::Imaging::SoftwareBitmap m_cachedSoftwareBitmap{ nullptr };

        // MISSING VARIABLE ADDED
        int32_t m_myProperty{ 0 };
    };
}

namespace winrt::PassportTool::factory_implementation
{
    struct MainWindow : MainWindowT<MainWindow, implementation::MainWindow>
    {
    };
}