#pragma once

#include <Windows.h>

#ifdef ASAUDIODLL_EXPORTS
#define ASAUDIODLL_API __declspec(dllexport)
#else
#define ASAUDIODLL_API __declspec(dllimport)
#endif

// --- 主要測試函式 ---
// 這是整個 DLL 的主要進入點
extern "C" ASAUDIODLL_API BSTR MacroTest(HWND ownerHwnd, const char* filePath);

// --- FFT 相關的 C-Style API ---
// 這些是為了與外部程式 (可能是 C 或 C#) 溝通而保留的 C 函式介面

// 設定測試參數
extern "C" ASAUDIODLL_API void fft_set_test_parm(
    const char* szOutDevName, int WaveOutVolumeValue,
    const char* szInDevName, int WaveInVolumeValue,
    int frequencyL, int frequencyR,
    int WaveOutdelay);

// 獲取緩衝區大小
extern "C" ASAUDIODLL_API int fft_get_buffer_size();

// 僅設定音量
extern "C" ASAUDIODLL_API void fft_set_audio_volume(
    const char* szOutDevName, int WaveOutVolumeValue,

    const char* szInDevName, int WaveInVolumeValue);

// 根據錯誤碼獲取錯誤訊息
extern "C" ASAUDIODLL_API BSTR fft_get_error_msg(int error_code);

// 設定系統音效靜音
extern "C" ASAUDIODLL_API void fft_set_system_sound_mute(bool Mute);

// 獲取 THD+N 和 dB 值
extern "C" ASAUDIODLL_API void fft_get_thd_n_db(double* thd_n, double* dB_ValueMax, double* freq);

// 獲取靜音時的 dB 值
extern "C" ASAUDIODLL_API void fft_get_mute_db(double* dB_ValueMax);

// 獲取信噪比 (SNR)
extern "C" ASAUDIODLL_API void fft_get_snr(double* snr);

// 執行 THD+N 測試
extern "C" ASAUDIODLL_API int fft_thd_n_exe(
    short* leftAudioData, short* rightAudioData,
    double* leftSpectrumData, double* rightSpectrumData);

// 執行靜音測試
extern "C" ASAUDIODLL_API int fft_mute_exe(
    bool MuteWaveOut, bool MuteWaveIn,
    short* leftAudioData, short* rightAudioData,
    double* leftSpectrumData, double* rightSpectrumData);