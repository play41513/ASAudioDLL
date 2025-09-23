#pragma once

#include <Windows.h>

#ifdef ASAUDIODLL_EXPORTS
#define ASAUDIODLL_API __declspec(dllexport)
#else
#define ASAUDIODLL_API __declspec(dllimport)
#endif

// --- �D�n���ը禡 ---
// �o�O��� DLL ���D�n�i�J�I
extern "C" ASAUDIODLL_API BSTR MacroTest(HWND ownerHwnd, const char* filePath);

// --- FFT ������ C-Style API ---
// �o�ǬO���F�P�~���{�� (�i��O C �� C#) ���q�ӫO�d�� C �禡����

// �]�w���հѼ�
extern "C" ASAUDIODLL_API void fft_set_test_parm(
    const char* szOutDevName, int WaveOutVolumeValue,
    const char* szInDevName, int WaveInVolumeValue,
    int frequencyL, int frequencyR,
    int WaveOutdelay);

// ����w�İϤj�p
extern "C" ASAUDIODLL_API int fft_get_buffer_size();

// �ȳ]�w���q
extern "C" ASAUDIODLL_API void fft_set_audio_volume(
    const char* szOutDevName, int WaveOutVolumeValue,

    const char* szInDevName, int WaveInVolumeValue);

// �ھڿ��~�X������~�T��
extern "C" ASAUDIODLL_API BSTR fft_get_error_msg(int error_code);

// �]�w�t�έ����R��
extern "C" ASAUDIODLL_API void fft_set_system_sound_mute(bool Mute);

// ��� THD+N �M dB ��
extern "C" ASAUDIODLL_API void fft_get_thd_n_db(double* thd_n, double* dB_ValueMax, double* freq);

// ����R���ɪ� dB ��
extern "C" ASAUDIODLL_API void fft_get_mute_db(double* dB_ValueMax);

// ����H���� (SNR)
extern "C" ASAUDIODLL_API void fft_get_snr(double* snr);

// ���� THD+N ����
extern "C" ASAUDIODLL_API int fft_thd_n_exe(
    short* leftAudioData, short* rightAudioData,
    double* leftSpectrumData, double* rightSpectrumData);

// �����R������
extern "C" ASAUDIODLL_API int fft_mute_exe(
    bool MuteWaveOut, bool MuteWaveIn,
    short* leftAudioData, short* rightAudioData,
    double* leftSpectrumData, double* rightSpectrumData);