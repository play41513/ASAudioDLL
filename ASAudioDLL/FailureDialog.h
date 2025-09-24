// FailureDialog.h
#pragma once
#include <Windows.h>
#include "ConfigManager.h" 
// 透過這個結構將音訊資料傳遞給 Dialog
struct AudioData {
    // 圖表原始資料
    short* leftAudioData;
    short* rightAudioData;
    double* leftSpectrumData;
    double* rightSpectrumData;
    int bufferSize;
    LPCWSTR errorMessage;

    // 新增：測試設定與結果
    const Config* config;         // 指向測試設定的指標
    double* thd_n_result;         // 指向實際測得的 THD+N 陣列 [L, R]
    double* FundamentalLevel_dBFS_result;   // 指向實際測得的 FundamentalLevel_dBFS 陣列 [L, R]
    double* freq_result;          // 指向實際測得的主頻率陣列 [L, R]
};

// 顯示 Dialog 的函式，如果使用者點擊 "Retry" 則回傳 true
bool ShowFailureDialog(HINSTANCE hInst, HWND ownerHwnd, AudioData* pData);