#define WIN32_LEAN_AND_MEAN
#define _CRT_SECURE_NO_WARNINGS
#include <ctime>
#include <windows.h>
#include <CommCtrl.h>
#include <shellapi.h>
#include <filesystem>
#include <fstream>
#include <chrono>
#include <iomanip>
#include <string>
#include <sstream>
#include <thread>

#pragma comment(lib, "comctl32.lib")
#pragma comment(linker, "\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

namespace fs = std::filesystem;

// Global variables
HWND hWnd, hEdit, hButton, hProgress, hLog, hOutputPath, hOpenButton;
std::wstring outputPath;

struct HumanReadable
{
    std::uintmax_t size{};

private:
    friend std::ostream& operator<<(std::ostream& os, HumanReadable hr)
    {
        int o{};
        double mantissa = hr.size;
        for (; mantissa >= 1024.; mantissa /= 1024., ++o);
        os << std::ceil(mantissa * 10.) / 10. << "BKMGTPE"[o];
        return o ? os << "B (" << hr.size << ')' : os;
    }
};

class CSVWriter {
private:
    std::ofstream file;
    size_t totalEntries = 0;
    size_t processedEntries = 0;

public:
    CSVWriter(const std::wstring& filename) : file(filename) {
        if (!file.is_open()) {
            throw std::runtime_error("Unable to open output file");
        }
        file << "Path,Name,Extension,Size, Last Update Date,Is Directory\n";
    }

    void writeEntry(const fs::directory_entry& entry) {
        auto path = entry.path();
        auto lastWriteTime = fs::last_write_time(entry);
        auto sctp = std::chrono::time_point_cast<std::chrono::system_clock::duration>(
            lastWriteTime - fs::file_time_type::clock::now() + std::chrono::system_clock::now());
        std::time_t cftime = std::chrono::system_clock::to_time_t(sctp);

        struct tm timeinfo;
        localtime_s(&timeinfo, &cftime);

        char timebuf[26];
        strftime(timebuf, sizeof(timebuf), "%Y-%m-%d %H:%M:%S", &timeinfo);

        file << path.string() << ","
            << path.filename().string() << ","
            << path.extension().string() << ","
            << HumanReadable{ fs::file_size(path) } << ","
            << timebuf << ","
            << (entry.is_directory() ? "Yes" : "No") << "\n";

        processedEntries++;
        updateProgress();
    }

    void setTotalEntries(size_t total) {
        totalEntries = total;
    }

    void updateProgress() {
        if (totalEntries > 0) {
            int percentage = static_cast<int>((processedEntries * 100) / totalEntries);
            SendMessage(hProgress, PBM_SETPOS, percentage, 0);
        }
    }

    ~CSVWriter() {
        if (file.is_open()) {
            file.close();
        }
    }
};

void appendLog(const std::wstring& message) {
    SendMessage(hLog, EM_SETSEL, (WPARAM)-1, (LPARAM)-1);
    SendMessage(hLog, EM_REPLACESEL, FALSE, (LPARAM)message.c_str());
    SendMessage(hLog, EM_SCROLLCARET, 0, 0);
}

size_t countEntries(const fs::path& dir) {
    size_t count = 0;
    for (auto const& entry : fs::recursive_directory_iterator(dir, fs::directory_options::skip_permission_denied)) {
        count++;
    }
    return count;
}

void processDirectory(const fs::path& dir, CSVWriter& writer) {
    try {
        for (const auto& entry : fs::recursive_directory_iterator(dir, fs::directory_options::skip_permission_denied)) {
            writer.writeEntry(entry);
            appendLog(L"Processed: " + entry.path().wstring() + L"\r\n");
        }
    }
    catch (const fs::filesystem_error& e) {
        std::wstringstream wss;
        wss << L"Error accessing path: " << e.path1().wstring() << L" - " << e.what() << L"\r\n";
        appendLog(wss.str());
    }
}

void runProgram() {
    wchar_t buffer[MAX_PATH];
    GetWindowText(hEdit, buffer, MAX_PATH);
    std::wstring directoryPath(buffer);

    try {
        WCHAR path[MAX_PATH];
        GetModuleFileName(NULL, path, MAX_PATH);
        std::wstring exePath(path);
        std::wstring dirPath = exePath.substr(0, exePath.find_last_of(L"\\/"));
        outputPath = dirPath + L"\\file_info.csv";

        CSVWriter writer(outputPath);

        appendLog(L"Counting entries...\r\n");
        size_t totalEntries = countEntries(directoryPath);
        writer.setTotalEntries(totalEntries);

        appendLog(L"Processing directory...\r\n");
        processDirectory(directoryPath, writer);

        appendLog(L"File information has been written to " + outputPath + L"\r\n");
        SetWindowText(hOutputPath, outputPath.c_str());
        EnableWindow(hOpenButton, TRUE);
    }
    catch (const std::exception& e) {
        std::wstringstream wss;
        wss << L"An error occurred: " << e.what() << L"\r\n";
        appendLog(wss.str());
    }
}

void openFileLocation() {
    WCHAR path[MAX_PATH];
    GetModuleFileName(NULL, path, MAX_PATH);
    std::wstring exePath(path);
    std::wstring dirPath = exePath.substr(0, exePath.find_last_of(L"\\/"));
    std::wstring fullPath = dirPath + L"\\" + outputPath;

    std::wstring command = L"/select,\"" + fullPath + L"\"";

    // Debug output
    MessageBox(NULL, fullPath.c_str(), L"File Path", MB_OK);
    MessageBox(NULL, command.c_str(), L"Command", MB_OK);

    HINSTANCE result = ShellExecute(NULL, L"open", L"explorer.exe", command.c_str(), NULL, SW_SHOWNORMAL);
    if ((INT_PTR)result <= 32) {
        // ShellExecute failed
        WCHAR errorMsg[256];
        swprintf_s(errorMsg, L"ShellExecute failed with error code: %d", (INT_PTR)result);
        MessageBox(NULL, errorMsg, L"Error", MB_OK | MB_ICONERROR);
    }
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_CREATE:
        hEdit = CreateWindow(L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_BORDER, 10, 10, 300, 25, hWnd, NULL, NULL, NULL);
        hButton = CreateWindow(L"BUTTON", L"Run", WS_CHILD | WS_VISIBLE, 320, 10, 100, 25, hWnd, (HMENU)1, NULL, NULL);
        hProgress = CreateWindowEx(0, PROGRESS_CLASS, NULL, WS_CHILD | WS_VISIBLE, 10, 45, 410, 25, hWnd, NULL, NULL, NULL);
        hLog = CreateWindowEx(WS_EX_CLIENTEDGE, L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY, 10, 80, 410, 200, hWnd, NULL, NULL, NULL);
        hOutputPath = CreateWindow(L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_READONLY, 10, 290, 300, 25, hWnd, NULL, NULL, NULL);
        hOpenButton = CreateWindow(L"BUTTON", L"Open Location", WS_CHILD | WS_VISIBLE | WS_DISABLED, 320, 290, 100, 25, hWnd, (HMENU)2, NULL, NULL);
        break;
    case WM_COMMAND:
        if (LOWORD(wParam) == 1) {
            std::thread(runProgram).detach();
        }
        else if (LOWORD(wParam) == 2) {
            openFileLocation();
        }
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int nCmdShow) {
    INITCOMMONCONTROLSEX icex;
    icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
    icex.dwICC = ICC_PROGRESS_CLASS;
    InitCommonControlsEx(&icex);

    WNDCLASSEX wcex = { 0 };
    wcex.cbSize = sizeof(WNDCLASSEX);
    wcex.style = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc = WndProc;
    wcex.hInstance = hInstance;
    wcex.hCursor = LoadCursor(NULL, IDC_ARROW);
    wcex.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wcex.lpszClassName = L"DirectoryProcessorClass";

    if (!RegisterClassEx(&wcex)) {
        MessageBox(NULL, L"Window Registration Failed!", L"Error!", MB_ICONEXCLAMATION | MB_OK);
        return 0;
    }

    hWnd = CreateWindow(L"DirectoryProcessorClass", L"Directory Processor", WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, 450, 360, NULL, NULL, hInstance, NULL);

    if (!hWnd) {
        MessageBox(NULL, L"Window Creation Failed!", L"Error!", MB_ICONEXCLAMATION | MB_OK);
        return 0;
    }

    ShowWindow(hWnd, nCmdShow);
    UpdateWindow(hWnd);

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return (int)msg.wParam;
}