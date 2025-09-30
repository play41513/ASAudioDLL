// FailureDialog.cpp
#include "pch.h"
#include "FailureDialog.h"
#include "resource.h"
#include <vector>
#include <algorithm>
#include <string>
#include <sstream>


enum class ChartView {
    THDN,
    SNR
};
static ChartView g_currentView = ChartView::THDN;

// 將全域資料指標儲存起來，以便在 DialogProc 中存取
static AudioData* g_pAudioData = nullptr;

constexpr int WAVEFORM_DISPLAY_SAMPLES = 256; // 一次顯示 256 個取樣點
static int g_scrollPosLeft = 0;
static int g_scrollPosRight = 0;



void UpdateResultText(HWND hDlg, const Config* cfg, AudioData* pData)
{
    if (!cfg || !pData) return;

    std::wstringstream wss;

    // --- 第 1 行: 播放與錄音裝置 ---
    std::wstring outDeviceDisplayName = pData->actualOutDeviceName.empty() 
        ? std::wstring(cfg->outDeviceName.begin(), cfg->outDeviceName.end()) 
        : pData->actualOutDeviceName;
    
    std::wstring inDeviceDisplayName = pData->actualInDeviceName.empty()
        ? std::wstring(cfg->inDeviceName.begin(), cfg->inDeviceName.end())
        : pData->actualInDeviceName;

    wss << L"Playback: " << outDeviceDisplayName << L" (" << cfg->outDeviceVolume << L"%)"
        << L"     |     Recording: " << inDeviceDisplayName << L" (" << cfg->inDeviceVolume << L"%)";
    SetDlgItemTextW(hDlg, IDC_INFO_DEVICES, wss.str().c_str());
    wss.str(L""); // 清空

    // --- 第 2 行: 頻率 [設定值, 實際值] ---
    wss << L"Frequency [Set, Actual]  L: [" << cfg->frequencyL << L", " << static_cast<int>(pData->freq_result[0]) << L"] Hz"
        << L"    R: [" << cfg->frequencyR << L", " << static_cast<int>(pData->freq_result[1]) << L"] Hz";
    SetDlgItemTextW(hDlg, IDC_INFO_FREQ, wss.str().c_str());
    wss.str(L"");

    // --- 第 3 行: THD+N [參數, 實際值] ---
    wss << L"THD+N [Threshold, Actual]  L: [< " << cfg->thd_n << ", " << static_cast<int>(pData->thd_n_result[0]) << L"] dB"
        << L"    R: [< " << cfg->thd_n << ", " << static_cast<int>(pData->thd_n_result[1]) << L"] dB";
    SetDlgItemTextW(hDlg, IDC_INFO_THDN, wss.str().c_str());
    wss.str(L"");

    // --- 第 4 行: dB Max [參數, 實際值] ---
    wss << L"Fundamental Level [Threshold, Actual]  L: [> " << cfg->FundamentalLevel_dBFS << ", " << static_cast<int>(pData->FundamentalLevel_dBFS_result[0]) << L"] dBFS"
        << L"    R: [> " << cfg->FundamentalLevel_dBFS << ", " << static_cast<int>(pData->FundamentalLevel_dBFS_result[1]) << L"] dBFS";
    SetDlgItemTextW(hDlg, IDC_INFO_DB, wss.str().c_str());
    wss.str(L"");
    // --- 第 4 行: SNR [參數, 實際值] ---
    if (pData->snr_result) {
        wss << L"SNR [Threshold, Actual]  L: [>" << cfg->snrThreshold << ", " << static_cast<int>(pData->snr_result[0]) << L"] dB"
            << L"     R: [>" << cfg->snrThreshold << ", " << static_cast<int>(pData->snr_result[1]) << L"] dB";
        SetDlgItemTextW(hDlg, IDC_INFO_SNR, wss.str().c_str());
        wss.str(L"");
    }
}

void DrawWaveform(HDC hdc, const RECT& rc, short* audioData, int startSample, int numSamples) {
    // --- 1. 設定繪圖區域 ---
    // 在左側為 Y 軸標籤預留 35px 的空間
    RECT graphRc = rc;
    graphRc.left += 20;

    // --- 2. 繪製背景和座標軸 ---
    HBRUSH hBrush = CreateSolidBrush(RGB(0, 0, 0));
    FillRect(hdc, &rc, hBrush); // 填滿整個控制項背景
    DeleteObject(hBrush);

    long width = graphRc.right - graphRc.left;
    long height = graphRc.bottom - graphRc.top;
    long midY = graphRc.top + height / 2;

    // 建立灰色畫筆用於座標軸
    HPEN hAxisPen = CreatePen(PS_SOLID, 1, RGB(128, 128, 128));
    HPEN hOldPen = (HPEN)SelectObject(hdc, hAxisPen);

    // 畫 X 軸 (Y=0)
    MoveToEx(hdc, graphRc.left, midY, NULL);
    LineTo(hdc, graphRc.right, midY);

    // 畫 Y 軸
    MoveToEx(hdc, graphRc.left, graphRc.top, NULL);
    LineTo(hdc, graphRc.left, graphRc.bottom);

    // --- 3. 繪製 Y 軸標籤 ---
    HFONT hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
    SelectObject(hdc, hFont);
    SetTextColor(hdc, RGB(200, 200, 200)); // 淺灰色文字
    SetBkMode(hdc, TRANSPARENT);

    // 準備文字繪製區域
    RECT textRc;
    textRc.left = rc.left;
    textRc.right = graphRc.left - 5; // 在 Y 軸左邊 5px 處

    // 標示最大值
    textRc.top = graphRc.top;
    textRc.bottom = graphRc.top + 15;
    DrawTextW(hdc, L"1", -1, &textRc, DT_RIGHT | DT_SINGLELINE | DT_VCENTER);

    // 標示零點
    textRc.top = midY - 7;
    textRc.bottom = midY + 7;
    DrawTextW(hdc, L"0", -1, &textRc, DT_RIGHT | DT_SINGLELINE | DT_VCENTER);

    // 標示最小值
    textRc.top = graphRc.bottom - 15;
    textRc.bottom = graphRc.bottom;
    DrawTextW(hdc, L"-1", -1, &textRc, DT_RIGHT | DT_SINGLELINE | DT_VCENTER);

    // --- 4. 繪製波形 ---
    HPEN hWavePen = CreatePen(PS_SOLID, 1, RGB(100, 200, 255));
    SelectObject(hdc, hWavePen); // 切換為波形畫筆

    if (audioData != nullptr && width > 1 && numSamples > 0) {
        MoveToEx(hdc, graphRc.left, midY - (audioData[startSample] * (height / 2)) / 32768, NULL);
        for (int i = 0; i < numSamples; ++i) {
            // 確保索引不會超出範圍
            if (startSample + i >= g_pAudioData->bufferSize) break;

            long x = graphRc.left + (i * width) / numSamples;
            long y = midY - (audioData[startSample + i] * (height / 2)) / 32768;
            LineTo(hdc, x, y);
        }
    }

    // --- 5. 清理資源 ---
    SelectObject(hdc, hOldPen);
    DeleteObject(hAxisPen);
    DeleteObject(hWavePen);
}


void DrawSpectrum(HDC hdc, const RECT& rc, double* spectrumData, int bufferSize) {
    // --- 1. 設定繪圖區域 ---
    // 為 Y 軸標籤預留左側 30px，為 X 軸標籤預留下方 20px
    RECT graphRc = rc;
    graphRc.left += 30;
    graphRc.bottom -= 20;

    // --- 2. 繪製背景和座標軸 ---
    HBRUSH hBrush = CreateSolidBrush(RGB(0, 0, 0));
    FillRect(hdc, &rc, hBrush);
    DeleteObject(hBrush);

    HPEN hAxisPen = CreatePen(PS_SOLID, 1, RGB(128, 128, 128)); // 灰色座標軸
    HPEN hOldPen = (HPEN)SelectObject(hdc, hAxisPen);

    // 畫 Y 軸
    MoveToEx(hdc, graphRc.left, graphRc.top, NULL);
    LineTo(hdc, graphRc.left, graphRc.bottom);
    // 畫 X 軸 (底部)
    MoveToEx(hdc, graphRc.left, graphRc.bottom, NULL);
    LineTo(hdc, graphRc.right, graphRc.bottom);

    // --- 3. 建立並設定字體 ---
    // 建立一個較小的字體
    LOGFONT lf = { 0 };
    lf.lfHeight = 15; // 字體高度
    wcscpy_s(lf.lfFaceName, L"MS Shell Dlg");
    HFONT hSmallFont = CreateFontIndirect(&lf);
    HFONT hOldFont = (HFONT)SelectObject(hdc, hSmallFont);
    SetTextColor(hdc, RGB(200, 200, 200));
    SetBkMode(hdc, TRANSPARENT);

    // --- 4. 繪製 Y 軸標籤 ---
    RECT textRc;
    textRc.left = rc.left;
    textRc.right = graphRc.left - 5;
    // 標示最大振幅
    textRc.top = graphRc.top;
    textRc.bottom = graphRc.top + 15;
    DrawTextW(hdc, L"Max", -1, &textRc, DT_RIGHT | DT_SINGLELINE | DT_VCENTER);
    // 標示最小振幅
    textRc.top = graphRc.bottom - 15;
    textRc.bottom = graphRc.bottom;
    DrawTextW(hdc, L"Min", -1, &textRc, DT_RIGHT | DT_SINGLELINE | DT_VCENTER);

    // --- 5. 繪製 X 軸 (頻率) 標籤 ---
    const int sampleRate = 44100;
    const int freqs[] = { 1000, 2000,3000,4000,5000,6000,7000,8000,9000, 10000,11000,12000,13000,14000,15000,16000,17000,18000,19000, 20000,21500};
    const wchar_t* labels[] = { L"1", L"2", L"3", L"4", L"5",L"6", L"7", L"8", L"9", L"10",L"11", L"12", L"13", L"14", L"15", L"16", L"17", L"18", L"19", L"20",L" kHz"};
    long width = graphRc.right - graphRc.left;

    for (int i = 0; i < sizeof(freqs) / sizeof(freqs[0]); ++i) {
        int freq = freqs[i];
        long x = graphRc.left + (long long)freq * bufferSize / sampleRate * width / (bufferSize / 2);

        MoveToEx(hdc, x, graphRc.bottom, NULL);
        LineTo(hdc, x, graphRc.bottom + 5);

        textRc.left = x - 15;
        textRc.top = graphRc.bottom + 5;
        textRc.right = x + 15;
        textRc.bottom = rc.bottom;
        DrawTextW(hdc, labels[i], -1, &textRc, DT_CENTER | DT_SINGLELINE);
    }

    // --- 6. 繪製頻譜資料 ---
    HPEN hSpectrumPen = CreatePen(PS_SOLID, 1, RGB(255, 100, 100));
    SelectObject(hdc, hSpectrumPen); // 切換為頻譜畫筆

    if (spectrumData != nullptr && width > 0) {
        int displaySize = bufferSize / 2;
        if (displaySize <= 0) return;

        double maxVal = 0.0;
        for (int i = 0; i < displaySize; ++i) if (spectrumData[i] > maxVal) maxVal = spectrumData[i];
        if (maxVal == 0.0) maxVal = 1.0;

        long height = graphRc.bottom - graphRc.top;
        if (height <= 0) return;

        for (int i = 0; i < displaySize; ++i) {
            long x = graphRc.left + (i * width) / displaySize;
            long y = graphRc.bottom - (long)((spectrumData[i] / maxVal) * height);
            MoveToEx(hdc, x, graphRc.bottom, NULL);
            LineTo(hdc, x, y);
        }
    }

    // --- 7. 清理資源 ---
    SelectObject(hdc, hOldFont);   // 還原舊字體
    SelectObject(hdc, hOldPen);    // 還原舊畫筆
    DeleteObject(hSmallFont);      // 刪除建立的字體
    DeleteObject(hAxisPen);
    DeleteObject(hSpectrumPen);
}

// Dialog 處理程序 (核心修改處)
INT_PTR CALLBACK FailureDialogProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_INITDIALOG: {
        g_pAudioData = (AudioData*)lParam;
        if (g_pAudioData && g_pAudioData->errorMessage) {
            SetWindowTextW(hDlg, L"測試失敗 (Test Failed)");
        }

        g_currentView = ChartView::THDN; // 每次打開都重設為預設
        CheckRadioButton(hDlg, IDC_RADIO_THDN, IDC_RADIO_SNR, IDC_RADIO_THDN);

        // 初始化捲軸
        if (g_pAudioData && g_pAudioData->bufferSize > WAVEFORM_DISPLAY_SAMPLES) {
            SCROLLINFO si = { 0 };
            si.cbSize = sizeof(si);
            si.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
            si.nMin = 0;
            si.nMax = g_pAudioData->bufferSize - WAVEFORM_DISPLAY_SAMPLES;
            si.nPage = WAVEFORM_DISPLAY_SAMPLES / 10; // 頁面滾動單位
            si.nPos = 0;

            SetScrollInfo(GetDlgItem(hDlg, IDC_SCROLL_WAVE_LEFT), SB_CTL, &si, TRUE);
            SetScrollInfo(GetDlgItem(hDlg, IDC_SCROLL_WAVE_RIGHT), SB_CTL, &si, TRUE);
        }
        g_scrollPosLeft = 0;
        g_scrollPosRight = 0;
        UpdateResultText(hDlg, g_pAudioData->config, g_pAudioData);
        return (INT_PTR)TRUE;
    }
    case WM_COMMAND: {
        // --- 處理按鈕點擊 ---
        if (LOWORD(wParam) == IDC_RETRY_BUTTON) {
            EndDialog(hDlg, TRUE);
            return (INT_PTR)TRUE;
        }
        if (LOWORD(wParam) == IDCANCEL) {
            EndDialog(hDlg, FALSE);
            return (INT_PTR)TRUE;
        }

        // <<< 處理 Radio Button 點擊事件 >>>
        bool needsRedraw = false;
        switch (LOWORD(wParam)) {
        case IDC_RADIO_THDN:
            if (g_currentView != ChartView::THDN) {
                g_currentView = ChartView::THDN;
                needsRedraw = true;
            }
            break;
        case IDC_RADIO_SNR:
            if (g_currentView != ChartView::SNR) {
                g_currentView = ChartView::SNR;
                needsRedraw = true;
            }
            break;
        }

        if (needsRedraw) {
            // 強制重繪所有四個圖表控制項
            InvalidateRect(GetDlgItem(hDlg, IDC_WAVEFORM_LEFT), NULL, TRUE);
            InvalidateRect(GetDlgItem(hDlg, IDC_WAVEFORM_RIGHT), NULL, TRUE);
            InvalidateRect(GetDlgItem(hDlg, IDC_SPECTRUM_LEFT), NULL, TRUE);
            InvalidateRect(GetDlgItem(hDlg, IDC_SPECTRUM_RIGHT), NULL, TRUE);
        }
        break;
    }
    case WM_HSCROLL: { // 處理水平捲軸事件
        int scrollId = GetDlgCtrlID((HWND)lParam);
        int* pScrollPos = nullptr;
        int waveformId = 0;

        if (scrollId == IDC_SCROLL_WAVE_LEFT) {
            pScrollPos = &g_scrollPosLeft;
            waveformId = IDC_WAVEFORM_LEFT;
        }
        else if (scrollId == IDC_SCROLL_WAVE_RIGHT) {
            pScrollPos = &g_scrollPosRight;
            waveformId = IDC_WAVEFORM_RIGHT;
        }

        if (pScrollPos) {
            SCROLLINFO si = { 0 };
            si.cbSize = sizeof(si);
            si.fMask = SIF_ALL;
            GetScrollInfo((HWND)lParam, SB_CTL, &si);

            int newPos = *pScrollPos;
            switch (LOWORD(wParam)) {
            case SB_LINELEFT: newPos -= 16; break;
            case SB_LINERIGHT: newPos += 16; break;
            case SB_PAGELEFT: newPos -= (int)si.nPage; break;
            case SB_PAGERIGHT: newPos += (int)si.nPage; break;
            case SB_THUMBTRACK: newPos = si.nTrackPos; break;
            }

            if (newPos < si.nMin) newPos = si.nMin;
            if (newPos > si.nMax) newPos = si.nMax;

            if (newPos != *pScrollPos) {
                *pScrollPos = newPos;
                SetScrollPos((HWND)lParam, SB_CTL, newPos, TRUE);
                InvalidateRect(GetDlgItem(hDlg, waveformId), NULL, TRUE);
            }
        }
        return TRUE;
    }
    case WM_DRAWITEM: {
        LPDRAWITEMSTRUCT lpDrawItem = (LPDRAWITEMSTRUCT)lParam;
        if (!g_pAudioData) return TRUE;

        // <<< 新增：根據 g_currentView 選擇要繪製的資料 >>>
        short* pLeftWaveData = (g_currentView == ChartView::THDN) ? g_pAudioData->leftAudioData : g_pAudioData->leftAudioData_SNR;
        short* pRightWaveData = (g_currentView == ChartView::THDN) ? g_pAudioData->rightAudioData : g_pAudioData->rightAudioData_SNR;
        double* pLeftSpectrumData = (g_currentView == ChartView::THDN) ? g_pAudioData->leftSpectrumData : g_pAudioData->leftSpectrumData_SNR;
        double* pRightSpectrumData = (g_currentView == ChartView::THDN) ? g_pAudioData->rightSpectrumData : g_pAudioData->rightSpectrumData_SNR;

        switch (lpDrawItem->CtlID) {
        case IDC_WAVEFORM_LEFT:
            DrawWaveform(lpDrawItem->hDC, lpDrawItem->rcItem, pLeftWaveData, g_scrollPosLeft, WAVEFORM_DISPLAY_SAMPLES);
            return TRUE;
        case IDC_WAVEFORM_RIGHT:
            DrawWaveform(lpDrawItem->hDC, lpDrawItem->rcItem, pRightWaveData, g_scrollPosRight, WAVEFORM_DISPLAY_SAMPLES);
            return TRUE;
        case IDC_SPECTRUM_LEFT:
            DrawSpectrum(lpDrawItem->hDC, lpDrawItem->rcItem, pLeftSpectrumData, g_pAudioData->bufferSize);
            return TRUE;
        case IDC_SPECTRUM_RIGHT:
            DrawSpectrum(lpDrawItem->hDC, lpDrawItem->rcItem, pRightSpectrumData, g_pAudioData->bufferSize);
            return TRUE;
        }
        break;
    }
    case WM_CLOSE: {
        EndDialog(hDlg, FALSE);
        return (INT_PTR)TRUE;
    }
    }
    return (INT_PTR)FALSE;
}

// 外部呼叫的函式 (無變動)
bool ShowFailureDialog(HINSTANCE hInst, HWND ownerHwnd, AudioData* pData) {
    return DialogBoxParam(hInst, MAKEINTRESOURCE(IDD_FAIL_DIALOG), ownerHwnd, FailureDialogProc, (LPARAM)pData) == TRUE;
}