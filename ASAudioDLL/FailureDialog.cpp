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

// �N�����ƫ����x�s�_�ӡA�H�K�b DialogProc ���s��
static AudioData* g_pAudioData = nullptr;

constexpr int WAVEFORM_DISPLAY_SAMPLES = 256; // �@����� 256 �Ө����I
static int g_scrollPosLeft = 0;
static int g_scrollPosRight = 0;



void UpdateResultText(HWND hDlg, const Config* cfg, AudioData* pData)
{
    if (!cfg || !pData) return;

    std::wstringstream wss;

    // --- �� 1 ��: ����P�����˸m ---
    wss << L"Playback: " << std::wstring(cfg->outDeviceName.begin(), cfg->outDeviceName.end()) << L" (" << cfg->outDeviceVolume << L"%)"
        << L"    |    Recording: " << std::wstring(cfg->inDeviceName.begin(), cfg->inDeviceName.end()) << L" (" << cfg->inDeviceVolume << L"%)";
    SetDlgItemTextW(hDlg, IDC_INFO_DEVICES, wss.str().c_str());
    wss.str(L""); // �M��

    // --- �� 2 ��: �W�v [�]�w��, ��ڭ�] ---
    wss << L"Frequency [Set, Actual]  L: [" << cfg->frequencyL << L", " << static_cast<int>(pData->freq_result[0]) << L"] Hz"
        << L"    R: [" << cfg->frequencyR << L", " << static_cast<int>(pData->freq_result[1]) << L"] Hz";
    SetDlgItemTextW(hDlg, IDC_INFO_FREQ, wss.str().c_str());
    wss.str(L"");

    // --- �� 3 ��: THD+N [�Ѽ�, ��ڭ�] ---
    wss << L"THD+N [Param, Actual]  L: [<" << cfg->thd_n << ", " << static_cast<int>(pData->thd_n_result[0]) << L"] dB"
        << L"    R: [>" << cfg->thd_n << ", " << static_cast<int>(pData->thd_n_result[1]) << L"] dB";
    SetDlgItemTextW(hDlg, IDC_INFO_THDN, wss.str().c_str());
    wss.str(L"");

    // --- �� 4 ��: dB Max [�Ѽ�, ��ڭ�] ---
    wss << L"FundamentalLevel_dBFS [Param, Actual]  L: [>" << cfg->FundamentalLevel_dBFS << ", " << static_cast<int>(pData->FundamentalLevel_dBFS_result[0]) << L"] dBFS"
        << L"    R: [>" << cfg->FundamentalLevel_dBFS << ", " << static_cast<int>(pData->FundamentalLevel_dBFS_result[1]) << L"] dBFS";
    SetDlgItemTextW(hDlg, IDC_INFO_DB, wss.str().c_str());
}

void DrawWaveform(HDC hdc, const RECT& rc, short* audioData, int startSample, int numSamples) {
    // --- 1. �]�wø�ϰϰ� ---
    // �b������ Y �b���ҹw�d 35px ���Ŷ�
    RECT graphRc = rc;
    graphRc.left += 20;

    // --- 2. ø�s�I���M�y�жb ---
    HBRUSH hBrush = CreateSolidBrush(RGB(0, 0, 0));
    FillRect(hdc, &rc, hBrush); // �񺡾�ӱ���I��
    DeleteObject(hBrush);

    long width = graphRc.right - graphRc.left;
    long height = graphRc.bottom - graphRc.top;
    long midY = graphRc.top + height / 2;

    // �إߦǦ�e���Ω�y�жb
    HPEN hAxisPen = CreatePen(PS_SOLID, 1, RGB(128, 128, 128));
    HPEN hOldPen = (HPEN)SelectObject(hdc, hAxisPen);

    // �e X �b (Y=0)
    MoveToEx(hdc, graphRc.left, midY, NULL);
    LineTo(hdc, graphRc.right, midY);

    // �e Y �b
    MoveToEx(hdc, graphRc.left, graphRc.top, NULL);
    LineTo(hdc, graphRc.left, graphRc.bottom);

    // --- 3. ø�s Y �b���� ---
    HFONT hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
    SelectObject(hdc, hFont);
    SetTextColor(hdc, RGB(200, 200, 200)); // �L�Ǧ��r
    SetBkMode(hdc, TRANSPARENT);

    // �ǳƤ�rø�s�ϰ�
    RECT textRc;
    textRc.left = rc.left;
    textRc.right = graphRc.left - 5; // �b Y �b���� 5px �B

    // �Х̤ܳj��
    textRc.top = graphRc.top;
    textRc.bottom = graphRc.top + 15;
    DrawTextW(hdc, L"1", -1, &textRc, DT_RIGHT | DT_SINGLELINE | DT_VCENTER);

    // �Хܹs�I
    textRc.top = midY - 7;
    textRc.bottom = midY + 7;
    DrawTextW(hdc, L"0", -1, &textRc, DT_RIGHT | DT_SINGLELINE | DT_VCENTER);

    // �Х̤ܳp��
    textRc.top = graphRc.bottom - 15;
    textRc.bottom = graphRc.bottom;
    DrawTextW(hdc, L"-1", -1, &textRc, DT_RIGHT | DT_SINGLELINE | DT_VCENTER);

    // --- 4. ø�s�i�� ---
    HPEN hWavePen = CreatePen(PS_SOLID, 1, RGB(100, 200, 255));
    SelectObject(hdc, hWavePen); // �������i�εe��

    if (audioData != nullptr && width > 1 && numSamples > 0) {
        MoveToEx(hdc, graphRc.left, midY - (audioData[startSample] * (height / 2)) / 32768, NULL);
        for (int i = 0; i < numSamples; ++i) {
            // �T�O���ޤ��|�W�X�d��
            if (startSample + i >= g_pAudioData->bufferSize) break;

            long x = graphRc.left + (i * width) / numSamples;
            long y = midY - (audioData[startSample + i] * (height / 2)) / 32768;
            LineTo(hdc, x, y);
        }
    }

    // --- 5. �M�z�귽 ---
    SelectObject(hdc, hOldPen);
    DeleteObject(hAxisPen);
    DeleteObject(hWavePen);
}


void DrawSpectrum(HDC hdc, const RECT& rc, double* spectrumData, int bufferSize) {
    // --- 1. �]�wø�ϰϰ� ---
    // �� Y �b���ҹw�d���� 30px�A�� X �b���ҹw�d�U�� 20px
    RECT graphRc = rc;
    graphRc.left += 30;
    graphRc.bottom -= 20;

    // --- 2. ø�s�I���M�y�жb ---
    HBRUSH hBrush = CreateSolidBrush(RGB(0, 0, 0));
    FillRect(hdc, &rc, hBrush);
    DeleteObject(hBrush);

    HPEN hAxisPen = CreatePen(PS_SOLID, 1, RGB(128, 128, 128)); // �Ǧ�y�жb
    HPEN hOldPen = (HPEN)SelectObject(hdc, hAxisPen);

    // �e Y �b
    MoveToEx(hdc, graphRc.left, graphRc.top, NULL);
    LineTo(hdc, graphRc.left, graphRc.bottom);
    // �e X �b (����)
    MoveToEx(hdc, graphRc.left, graphRc.bottom, NULL);
    LineTo(hdc, graphRc.right, graphRc.bottom);

    // --- 3. �إߨó]�w�r�� ---
    // �إߤ@�Ӹ��p���r��
    LOGFONT lf = { 0 };
    lf.lfHeight = 15; // �r�鰪��
    wcscpy_s(lf.lfFaceName, L"MS Shell Dlg");
    HFONT hSmallFont = CreateFontIndirect(&lf);
    HFONT hOldFont = (HFONT)SelectObject(hdc, hSmallFont);
    SetTextColor(hdc, RGB(200, 200, 200));
    SetBkMode(hdc, TRANSPARENT);

    // --- 4. ø�s Y �b���� ---
    RECT textRc;
    textRc.left = rc.left;
    textRc.right = graphRc.left - 5;
    // �Х̤ܳj���T
    textRc.top = graphRc.top;
    textRc.bottom = graphRc.top + 15;
    DrawTextW(hdc, L"Max", -1, &textRc, DT_RIGHT | DT_SINGLELINE | DT_VCENTER);
    // �Х̤ܳp���T
    textRc.top = graphRc.bottom - 15;
    textRc.bottom = graphRc.bottom;
    DrawTextW(hdc, L"Min", -1, &textRc, DT_RIGHT | DT_SINGLELINE | DT_VCENTER);

    // --- 5. ø�s X �b (�W�v) ���� ---
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

    // --- 6. ø�s�W�и�� ---
    HPEN hSpectrumPen = CreatePen(PS_SOLID, 1, RGB(255, 100, 100));
    SelectObject(hdc, hSpectrumPen); // �������W�еe��

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

    // --- 7. �M�z�귽 ---
    SelectObject(hdc, hOldFont);   // �٭��¦r��
    SelectObject(hdc, hOldPen);    // �٭��µe��
    DeleteObject(hSmallFont);      // �R���إߪ��r��
    DeleteObject(hAxisPen);
    DeleteObject(hSpectrumPen);
}

// Dialog �B�z�{�� (�֤߭ק�B)
INT_PTR CALLBACK FailureDialogProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_INITDIALOG: {
        g_pAudioData = (AudioData*)lParam;
        if (g_pAudioData && g_pAudioData->errorMessage) {
            SetWindowTextW(hDlg, L"���ե��� (Test Failed)");
        }

        g_currentView = ChartView::THDN; // �C�����}�����]���w�]
        CheckRadioButton(hDlg, IDC_RADIO_THDN, IDC_RADIO_SNR, IDC_RADIO_THDN);

        // ��l�Ʊ��b
        if (g_pAudioData && g_pAudioData->bufferSize > WAVEFORM_DISPLAY_SAMPLES) {
            SCROLLINFO si = { 0 };
            si.cbSize = sizeof(si);
            si.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
            si.nMin = 0;
            si.nMax = g_pAudioData->bufferSize - WAVEFORM_DISPLAY_SAMPLES;
            si.nPage = WAVEFORM_DISPLAY_SAMPLES / 10; // �����u�ʳ��
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
        // --- �B�z���s�I�� ---
        if (LOWORD(wParam) == IDC_RETRY_BUTTON) {
            EndDialog(hDlg, TRUE);
            return (INT_PTR)TRUE;
        }
        if (LOWORD(wParam) == IDCANCEL) {
            EndDialog(hDlg, FALSE);
            return (INT_PTR)TRUE;
        }

        // <<< �B�z Radio Button �I���ƥ� >>>
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
            // �j�ø�Ҧ��|�ӹϪ��
            InvalidateRect(GetDlgItem(hDlg, IDC_WAVEFORM_LEFT), NULL, TRUE);
            InvalidateRect(GetDlgItem(hDlg, IDC_WAVEFORM_RIGHT), NULL, TRUE);
            InvalidateRect(GetDlgItem(hDlg, IDC_SPECTRUM_LEFT), NULL, TRUE);
            InvalidateRect(GetDlgItem(hDlg, IDC_SPECTRUM_RIGHT), NULL, TRUE);
        }
        break;
    }
    case WM_HSCROLL: { // �B�z�������b�ƥ�
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

        // <<< �s�W�G�ھ� g_currentView ��ܭnø�s����� >>>
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

// �~���I�s���禡 (�L�ܰ�)
bool ShowFailureDialog(HINSTANCE hInst, HWND ownerHwnd, AudioData* pData) {
    return DialogBoxParam(hInst, MAKEINTRESOURCE(IDD_FAIL_DIALOG), ownerHwnd, FailureDialogProc, (LPARAM)pData) == TRUE;
}