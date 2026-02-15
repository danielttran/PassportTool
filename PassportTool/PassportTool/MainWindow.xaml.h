#pragma once

#include "MainWindow.g.h"
#include <winrt/Windows.Graphics.Imaging.h>
#include <winrt/Microsoft.UI.Xaml.Media.Imaging.h>
#include <winrt/Microsoft.UI.Input.h>
#include <winrt/Windows.ApplicationModel.DataTransfer.h>
#include <winrt/Windows.Storage.h>

namespace winrt::PassportTool::implementation
{
    struct MainWindow : MainWindowT<MainWindow>
    {
        MainWindow();

        // Life Cycle
        void OnWindowLoaded(winrt::Windows::Foundation::IInspectable const& sender, winrt::Microsoft::UI::Xaml::RoutedEventArgs const& e);

        // UI Event Handlers
        winrt::fire_and_forget BtnPickImage_Click(winrt::Windows::Foundation::IInspectable const& sender, winrt::Microsoft::UI::Xaml::RoutedEventArgs const& e);
        winrt::fire_and_forget BtnApplyCrop_Click(winrt::Windows::Foundation::IInspectable const& sender, winrt::Microsoft::UI::Xaml::RoutedEventArgs const& e);
        winrt::fire_and_forget BtnSaveSheet_Click(winrt::Windows::Foundation::IInspectable const& sender, winrt::Microsoft::UI::Xaml::RoutedEventArgs const& e);
        void BtnRotate_Click(winrt::Windows::Foundation::IInspectable const& sender, winrt::Microsoft::UI::Xaml::RoutedEventArgs const& e);

        void OnUnitChanged(winrt::Windows::Foundation::IInspectable const& sender, winrt::Microsoft::UI::Xaml::RoutedEventArgs const& e);
        void OnSettingsChanged(winrt::Microsoft::UI::Xaml::Controls::NumberBox const& sender, winrt::Microsoft::UI::Xaml::Controls::NumberBoxValueChangedEventArgs const& args);

        // Input Handling
        void OnImagePointerPressed(winrt::Windows::Foundation::IInspectable const& sender, winrt::Microsoft::UI::Xaml::Input::PointerRoutedEventArgs const& e);
        void OnImagePointerMoved(winrt::Windows::Foundation::IInspectable const& sender, winrt::Microsoft::UI::Xaml::Input::PointerRoutedEventArgs const& e);
        void OnImagePointerReleased(winrt::Windows::Foundation::IInspectable const& sender, winrt::Microsoft::UI::Xaml::Input::PointerRoutedEventArgs const& e);
        void OnImagePointerWheelChanged(winrt::Windows::Foundation::IInspectable const& sender, winrt::Microsoft::UI::Xaml::Input::PointerRoutedEventArgs const& e);
        void OnCropSizeChanged(winrt::Windows::Foundation::IInspectable const& sender, winrt::Microsoft::UI::Xaml::SizeChangedEventArgs const& e);

        void OnDragOver(winrt::Windows::Foundation::IInspectable const& sender, winrt::Microsoft::UI::Xaml::DragEventArgs const& e);
        winrt::fire_and_forget OnDrop(winrt::Windows::Foundation::IInspectable const& sender, winrt::Microsoft::UI::Xaml::DragEventArgs const& e);

    private:
        // Internal Logic
        winrt::Windows::Foundation::IAsyncAction RegeneratePreviewGrid();
        void UpdateSheetSize();
        void RefitCropContainer();
        void UpdateCellDimensionsDisplay();
        double GetPixelsPerUnit();
        void Log(winrt::hstring const& message); // Logging Helper

        // High-res processing
        winrt::Windows::Foundation::IAsyncAction LoadImageFromFile(winrt::Windows::Storage::StorageFile file);
        winrt::Windows::Foundation::IAsyncOperation<winrt::Windows::Graphics::Imaging::SoftwareBitmap> CaptureCropAsBitmap();
        winrt::Windows::Foundation::IAsyncOperation<winrt::Windows::Graphics::Imaging::SoftwareBitmap> GenerateFullResolutionSheet();
        winrt::Windows::Foundation::IAsyncOperation<winrt::Windows::Graphics::Imaging::SoftwareBitmap> CaptureElementAsync(winrt::Microsoft::UI::Xaml::UIElement element);

        // State
        bool m_isLoaded{ false };
        winrt::Windows::Graphics::Imaging::SoftwareBitmap m_originalBitmap{ nullptr };
        winrt::Windows::Graphics::Imaging::SoftwareBitmap m_croppedStamp{ nullptr };

        bool m_isDragging{ false };
        bool m_zoomingFromMouse{ false };
        winrt::Windows::Foundation::Point m_lastPoint{ 0,0 };
    };
}

namespace winrt::PassportTool::factory_implementation
{
    struct MainWindow : MainWindowT<MainWindow, implementation::MainWindow>
    {
    };
}