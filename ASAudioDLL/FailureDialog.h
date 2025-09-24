// FailureDialog.h
#pragma once
#include <Windows.h>
#include "ConfigManager.h" 
// �z�L�o�ӵ��c�N���T��ƶǻ��� Dialog
struct AudioData {
    // �Ϫ��l���
    short* leftAudioData;
    short* rightAudioData;
    double* leftSpectrumData;
    double* rightSpectrumData;
    int bufferSize;
    LPCWSTR errorMessage;

    // �s�W�G���ճ]�w�P���G
    const Config* config;         // ���V���ճ]�w������
    double* thd_n_result;         // ���V��ڴ��o�� THD+N �}�C [L, R]
    double* FundamentalLevel_dBFS_result;   // ���V��ڴ��o�� FundamentalLevel_dBFS �}�C [L, R]
    double* freq_result;          // ���V��ڴ��o���D�W�v�}�C [L, R]
};

// ��� Dialog ���禡�A�p�G�ϥΪ��I�� "Retry" �h�^�� true
bool ShowFailureDialog(HINSTANCE hInst, HWND ownerHwnd, AudioData* pData);