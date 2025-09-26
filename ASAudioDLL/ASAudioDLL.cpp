#include "ConfigManager.h"
#include "AudioManager.h"
#include "ASAudioDLL.h"
#include "ASAudio.h"
#include "FailureDialog.h" 
#include "resource.h"     
#include <sstream>
#include <fstream>
#include <algorithm>
#include <thread>
#include <chrono>
#include <locale>
#include <codecvt>

#include <mmsystem.h>
#include <mmdeviceapi.h>
#include <endpointvolume.h>
#include <dshow.h>
#include <propkey.h>
#include <functiondiscoverykeys.h>
#include <atlbase.h> // for CComPtr
#include <initguid.h>
#include <comdef.h> // for _bstr_t

short leftAudioData[BUFFER_SIZE] = { 0 };
short rightAudioData[BUFFER_SIZE] = { 0 };
double leftSpectrumData[BUFFER_SIZE] = { 0 };
double rightSpectrumData[BUFFER_SIZE] = { 0 };

short leftAudioData_SNR[BUFFER_SIZE] = { 0 };
short rightAudioData_SNR[BUFFER_SIZE] = { 0 };
double leftSpectrumData_SNR[BUFFER_SIZE] = { 0 };
double rightSpectrumData_SNR[BUFFER_SIZE] = { 0 };

double thd_n[2] = { 0 };
double FundamentalLevel_dBFS[2] = { 0 };
double freq[2] = { 0 };

extern HINSTANCE g_hInst;

extern "C" __declspec(dllexport) BSTR MacroTest(HWND ownerHwnd, const char* filePath)
{
	bool shouldRetry;
	std::string finalResult;
	ASAudio& audioInstance = ASAudio::GetInstance(); // ������

	do {
		shouldRetry = false;
		audioInstance.ClearLog();

		Config configs;
		if (!ConfigManager::ReadConfig(filePath, configs)) {
			finalResult = "LOG:ERROR_CONFIG_NOT_FIND#";
		}
		else {
			// ��l�ƴ��հѼơA�z�L ASAudio ��ҳ]�w
			fft_set_test_parm(
				configs.outDeviceName.c_str(), configs.outDeviceVolume,
				configs.inDeviceName.c_str(), configs.inDeviceVolume,
				configs.frequencyL, configs.frequencyR,
				configs.waveOutDelay
			);
			audioInstance.GetWaveParm().AudioFile = configs.WAVFilePath_w;
			audioInstance.GetWaveParm().fundamentalBandwidthHz = configs.fundamentalBandwidthHz;

			AudioManager audioManager;
			audioManager.ExecuteTestsFromConfig(configs);
			finalResult = audioManager.GetResultString(); // �`�O������G�A�L�צ��\�Υ���
		}

		bool hasError = (finalResult.find("ERROR") != std::string::npos);
		if (hasError && ownerHwnd != NULL) {
			/*
			std::string strTemp = "Failed: " + finalResult;
			std::wstring wErrorMsg = audioInstance.charToWstring(strTemp.c_str()) + L"\n\n" + audioInstance.GetLog() + L"\n\n�O�_�n�����H(Retry ?)";
			if (MessageBoxW(ownerHwnd, wErrorMsg.c_str(), L"���ե���", MB_YESNO | MB_ICONERROR) == IDYES) {
				shouldRetry = true;
			}*/
			std::string strTemp = "���� (Failed): " + finalResult;
			std::wstring wErrorMsg = audioInstance.charToWstring(strTemp.c_str()) + L"\n" + audioInstance.GetLog();

			AudioData data;
			data.bufferSize = BUFFER_SIZE;
			data.errorMessage = wErrorMsg.c_str();
			// �ǤJ�]�w�M���G������
			data.config = &configs;
			data.thd_n_result = thd_n;
			data.FundamentalLevel_dBFS_result = FundamentalLevel_dBFS;
			data.freq_result = freq;

			// ��R THD+N ���Ϫ���
			data.leftAudioData = leftAudioData;
			data.rightAudioData = rightAudioData;
			data.leftSpectrumData = leftSpectrumData;
			data.rightSpectrumData = rightSpectrumData;

			// <<< �s�W�G��R SNR ���Ϫ��� >>>
			data.leftAudioData_SNR = leftAudioData_SNR;
			data.rightAudioData_SNR = rightAudioData_SNR;
			data.leftSpectrumData_SNR = leftSpectrumData_SNR;
			data.rightSpectrumData_SNR = rightSpectrumData_SNR;

			// �I�s ShowFailureDialog�A�{�b���֦��F��ո��
			if (ShowFailureDialog(g_hInst, ownerHwnd, &data)) {
				shouldRetry = true;
			}
		}

	} while (shouldRetry);

	if (finalResult.empty()) {
		finalResult = "LOG:PASS#";
	}

	// �N string �ର BSTR
	std::wstring wstr = audioInstance.charToWstring(finalResult.c_str());
	return SysAllocString(wstr.c_str());
}



void __cdecl fft_set_test_parm(
	const char* szOutDevName, int WaveOutVolumeValue
	, const char* szInDevName, int WaveInVolumeValue
	, int frequencyL, int frequencyR
	, int WaveOutdelay)
{
	auto& parm = ASAudio::GetInstance().GetWaveParm();
	parm.WaveOutDev = ASAudio::GetInstance().charToWstring(szOutDevName);
	parm.WaveInDev = ASAudio::GetInstance().charToWstring(szInDevName);
	parm.WaveOutVolume = WaveOutVolumeValue;
	parm.WaveInVolume = WaveInVolumeValue;
	parm.frequencyL = frequencyL;
	parm.frequencyR = frequencyR;
	parm.WaveOutDelay = WaveOutdelay;
}

int __cdecl fft_get_buffer_size()
{
	return BUFFER_SIZE;
}

void __cdecl fft_set_audio_volume(const char* szOutDevName, int WaveOutVolumeValue
	, const char* szInDevName, int WaveInVolumeValue)
{
	// ��� ASAudio ����� (singleton)
	ASAudio& audioInstance = ASAudio::GetInstance();

	// �z�L��Ҧs���� public ���� mWAVE_PARM
	auto& parm = audioInstance.GetWaveParm();

	// �N�ǤJ�� C-style �r���ഫ�� wstring �ó]�w�Ѽ�
	parm.WaveOutDev = audioInstance.charToWstring(szOutDevName);
	parm.WaveInDev = audioInstance.charToWstring(szInDevName);
	parm.WaveOutVolume = WaveOutVolumeValue;
	parm.WaveInVolume = WaveInVolumeValue;

	// �I�s ASAudio �������禡�ӹ�ڮM�γ��J���M��z�����q
	audioInstance.SetMicSystemVolume();
	audioInstance.SetSpeakerSystemVolume();
}

BSTR __cdecl fft_get_error_msg(int error_code)
{
	ErrorCode code = static_cast<ErrorCode>(error_code);
	auto it = error_map.find(code);
	if (it != error_map.end()) {
		return SysAllocString(it->second);
	}
	else {
		return SysAllocString(L"Unknown error code");
	}
}

void __cdecl fft_set_system_sound_mute(bool Mute)
{ // �t�ΰѼƳ]�m�ASPIF_SENDCHANGE ��ܥߧY���ASPIF_UPDATEINIFILE ��ܧ�s�t�m���A�ĤG�ӰѼ�1�B0����windosw�w�]�B�L����
	int flag = Mute ? 0 : 1;
	SystemParametersInfo(SPI_SETSOUNDSENTRY, flag, NULL, SPIF_SENDCHANGE | SPIF_UPDATEINIFILE);
}

void __cdecl fft_get_thd_n_db(double* thd_n, double* FundamentalLevel_dBFS, double* freq)
{
	auto& parm = ASAudio::GetInstance().GetWaveParm();
	thd_n[0] = parm.leftWAVE_ANALYSIS.thd_N_dB;
	thd_n[1] = parm.rightWAVE_ANALYSIS.thd_N_dB;

	// 1. �q��q�]���T������^�}�ڸ��A�٭�^���T
	double leftAmplitude = sqrt(parm.leftWAVE_ANALYSIS.fundamentalEnergy);
	double rightAmplitude = sqrt(parm.rightWAVE_ANALYSIS.fundamentalEnergy);

	// 2. �H 16-bit ���̤j�� 32767.0 �@���ѦҡA�p�� dBFS
	//    �ϥ� 20 * log10() ���ഫ���T
	double fullScaleReference = 32767.0;

	if (leftAmplitude > 0) {
		FundamentalLevel_dBFS[0] = 20 * log10(leftAmplitude / fullScaleReference);
	}
	else {
		FundamentalLevel_dBFS[0] = -INFINITY; // �R��
	}

	if (rightAmplitude > 0) {
		FundamentalLevel_dBFS[1] = 20 * log10(rightAmplitude / fullScaleReference);
	}
	else {
		FundamentalLevel_dBFS[1] = -INFINITY; // �R��
	}

	// �W�v
	const double sampleRate = 44100.0;
	const int N = BUFFER_SIZE;
	freq[0] = parm.leftWAVE_ANALYSIS.TotalEnergyPoint * (sampleRate / N);
	freq[1] = parm.rightWAVE_ANALYSIS.TotalEnergyPoint * (sampleRate / N);
	/*
	THD+N��dB�Ƚd��
	�M�~���T�]�ơG
	���~�誺�M�~���T�]�Ƴq�`�n�DTHD+N�C�� -80 dB�A�o�N���ۥ��u�M���n�D�`�C�C
	���O�q�l���~�G
	�a�έ��T�M���O�q�l���~��THD+N�@��b -60 dB �� -80 dB �����C
	�����~�譵�T�]�ơG
	THD+N�Ȧb -60 dB �� -40 dB �����Q�{���O�����~��C
	�C�~�譵�T�]�ơG
	THD+N�Ȱ��� -40 dB �i���ܥ��u�M���n�����A������t�C
	�̤jdB�Ƚd��
	�M�~���T�]�ơG
	�M�~�]�ƪ��̤jdB�ȡ]��ܫH���j�ס^�q�`�b 80 dB �� 120 dB �����C
	���O�q�l���~�G
	�a�έ��T�M���O�q�l���~���̤jdB�ȳq�`�b 60 dB �� 100 dB �����C
	�����~�譵�T�]�ơG
	�̤jdB�Ȧb 70 dB �� 90 dB �����C
	�C�~�譵�T�]�ơG
	�̤jdB�ȧC�� 70 dB�C
	*/
}

void __cdecl fft_get_mute_db(double* FundamentalLevel_dBFS)
{
	auto& parm = ASAudio::GetInstance().GetWaveParm();
	FundamentalLevel_dBFS[0] = 10 * log10(parm.leftMuteWAVE_ANALYSIS.TotalEnergy);
	FundamentalLevel_dBFS[1] = 10 * log10(parm.rightMuteWAVE_ANALYSIS.TotalEnergy);
	/*
	�R�����ժ��̤jdB�Ƚd��
	�M�~���T�]�ơG
	���~�誺�M�~���T�]�Ʀb�R�����A�U�A�̤jdB�ȡ]�I�����n�q���^�q�`�C�� -90 dB�C
	���O�q�l���~�G
	�a�έ��T�M���O�q�l���~���R���̤jdB�Ȥ@��b -70 dB �� -90 dB �����C
	�����~�譵�T�]�ơG
	�R���̤jdB�Ȧb -60 dB �� -80 dB �����C
	�C�~�譵�T�]�ơG
	�R���̤jdB�Ȱ��� -60 dB ��ܦb�R�����A�U���n���������C
	*/
}

void __cdecl fft_get_snr(double* snr)
{
	auto& parm = ASAudio::GetInstance().GetWaveParm();

	// �p���q��
	double left_snr_ratio = (parm.leftMuteWAVE_ANALYSIS.TotalEnergy > 0) ? (parm.leftWAVE_ANALYSIS.fundamentalEnergy / parm.leftMuteWAVE_ANALYSIS.TotalEnergy) : 0;
	double right_snr_ratio = (parm.rightMuteWAVE_ANALYSIS.TotalEnergy > 0) ? (parm.rightWAVE_ANALYSIS.fundamentalEnergy / parm.rightMuteWAVE_ANALYSIS.TotalEnergy) : 0;

	// �B�z���n���s�����ݱ��p ---
	const double MAX_SNR_CAP = 144.0; // �]�w�@�ӫD�`���� SNR �W�� (dB)�A�N���n�����G

	if (left_snr_ratio > 0) {
		snr[0] = 10 * log10(left_snr_ratio);
	}
	else {
		// �p�G��q�� 0 (�]�����n�� 0)�A�����ᤩ�W����
		snr[0] = MAX_SNR_CAP;
	}

	if (right_snr_ratio > 0) {
		snr[1] = 10 * log10(right_snr_ratio);
	}
	else {
		snr[1] = MAX_SNR_CAP;
	}

	// �i�H��ܩʦa�勵�`�p��X���Ȥ]�]�w�W���A�קK�X�{�L�j���Ʀr
	if (snr[0] > MAX_SNR_CAP) snr[0] = MAX_SNR_CAP;
	if (snr[1] > MAX_SNR_CAP) snr[1] = MAX_SNR_CAP;
	/*
	SNR ���ժ��d��
	�M�~���T�]�ơG
	���~�誺�M�~���T�]�Ƴq�`�㦳�D�`���� SNR�A�d��b 90 dB �� 120 dB �����A���ǰ��ݳ]�ƬƦܥi�H�F�� 120 dB �H�W�C
	���O�q�l���~�G
	�a�έ��T�M���O�q�l���~�� SNR �d��q�`�b 70 dB �� 100 dB �����C
	�����~�譵�T�]�ơG
	�����~�譵�T�]�ƪ� SNR �q�`�b 60 dB �� 80 dB �����C
	�C�~�譵�T�]�ơG
	�C�~�譵�T�]�ƪ� SNR �i��C�� 60 dB�A��ܫH���P���n�������t�Z���p�A����i����t�C
	SNR ���ժ�����
	�� SNR�G
	�� SNR �Ȫ�ܭI�����n�C�A���W�H���M���B�j�j�C�o�b�M�~���T�t�Τ��׬����n�A�]�����̭n�D���ײM��������M�C���n�C
	�C SNR�G
	�C SNR �Ȫ�ܭI�����n�۹�����A�H�������M���C�b�@�Ǿ��n�e�ԫ׸��������Τ��A�o�i�ण�|�y����ۼv�T�A���b���n�D�����W���Τ��A�C SNR �i��|��ۭ��C����C
	*/
}

// �~��C��ơA�Ω����FFT�ഫ
int __cdecl fft_thd_n_exe(short* leftAudioData, short* rightAudioData, double* leftSpectrumData, double* rightSpectrumData)
{
	auto& parm = ASAudio::GetInstance().GetWaveParm();
	parm.bMuteTest = false;
	MMRESULT result;

	result = ASAudio::GetInstance().SetMicSystemVolume();
	result = ASAudio::GetInstance().SetSpeakerSystemVolume();
	if (result != MMSYSERR_NOERROR)
		return result;

	result = ASAudio::GetInstance().StartRecordingAndDrawSpectrum();
	if (result != MMSYSERR_NOERROR)
		return result;

	WAVE_PARM::RecordingResult waitResult = ASAudio::GetInstance().WaitForRecording(10);

	if (waitResult != WAVE_PARM::RecordingResult::Success) {
		ASAudio::GetInstance().StopRecording(); 
		ASAudio::GetInstance().stopPlayback();

		if (waitResult == WAVE_PARM::RecordingResult::Timeout) {
			return ERROR_CODE_UNSET_PARAMETER; 
		}
		else if (waitResult == WAVE_PARM::RecordingResult::DeviceError) {
			return ERROR_CODE_UNSET_PARAMETER; 
		}
	}

	ASAudio::GetInstance().StopRecording(); // �������
	ASAudio::GetInstance().stopPlayback(); // �������

	const fftw_complex* leftSpectrum = ASAudio::GetInstance().GetLeftSpectrumPtr();
	const fftw_complex* rightSpectrum = ASAudio::GetInstance().GetRightSpectrumPtr();

	memcpy(leftAudioData, parm.WAVE_DATA.LeftAudioBuffer, BUFFER_SIZE * sizeof(short));
	memcpy(rightAudioData, parm.WAVE_DATA.RightAudioBuffer, BUFFER_SIZE * sizeof(short));

	double maxAmplitude = 0;
	for (int i = 0; i < BUFFER_SIZE; ++i) {
		double sample = static_cast<double>(leftAudioData[i]);
		if (fabs(sample) > maxAmplitude) {
			maxAmplitude = fabs(sample);
		}
	}
	maxAmplitude = max(maxAmplitude, 1.0); // �קK log10(0) �����p
	double maxAmplitude_dB = 20 * log10(maxAmplitude / 32767);

	for (int i = 0; i < BUFFER_SIZE; ++i) {
		leftSpectrumData[i] = sqrt(leftSpectrum[i][0] * leftSpectrum[i][0] + leftSpectrum[i][1] * leftSpectrum[i][1]) / BUFFER_SIZE;
		rightSpectrumData[i] = sqrt(rightSpectrum[i][0] * rightSpectrum[i][0] + rightSpectrum[i][1] * rightSpectrum[i][1]) / BUFFER_SIZE;
	}

	// ���R�W��THD+N
	ASAudio::GetInstance().SpectrumAnalysis(leftSpectrumData, rightSpectrumData);
	return result;
}

// �~��C��ơA�Ω����FFT�ഫ
int __cdecl fft_mute_exe(bool MuteWaveOut, bool MuteWaveIn,
	short* leftAudioData, short* rightAudioData, double* leftSpectrumData, double* rightSpectrumData)
{
	auto& parm = ASAudio::GetInstance().GetWaveParm();
	parm.bMuteTest = true;
	MMRESULT result = MMSYSERR_NOERROR;

	if (MuteWaveOut) {
		// �T�O����i��b�����n��������
		ASAudio::GetInstance().stopPlayback();
	}

	if (MuteWaveIn)
		ASAudio::GetInstance().SetMicMute(true);
	result = ASAudio::GetInstance().StartRecordingOnly();

	if (result != MMSYSERR_NOERROR)
	{
		parm.bMuteTest = false;
		if (MuteWaveIn) 
			ASAudio::GetInstance().SetMicMute(false);
		return result;
	}

	WAVE_PARM::RecordingResult waitResult = ASAudio::GetInstance().WaitForRecording(10);
	if (waitResult != WAVE_PARM::RecordingResult::Success) {
		ASAudio::GetInstance().StopRecording(); 
		ASAudio::GetInstance().stopPlayback();

		// Restore state even on error
		parm.bMuteTest = false;
		if (MuteWaveIn) ASAudio::GetInstance().SetMicMute(false);
		if (MuteWaveOut) ASAudio::GetInstance().SetSpeakerSystemVolume();

		if (waitResult == WAVE_PARM::RecordingResult::Timeout) {
			return ERROR_CODE_UNSET_PARAMETER; 
		}
		else if (waitResult == WAVE_PARM::RecordingResult::DeviceError) {
			return ERROR_CODE_UNSET_PARAMETER; 
		}
	}

	ASAudio::GetInstance().StopRecording();// �������
	ASAudio::GetInstance().stopPlayback(); // �������

	ASAudio::GetInstance().ExecuteFft();

	fftw_complex* leftSpectrum;
	fftw_complex* rightSpectrum;
	ASAudio::GetInstance().GetSpectrumData(leftSpectrum, rightSpectrum); // ����W�мƾ�
	memcpy(leftAudioData, parm.WAVE_DATA.LeftAudioBuffer, BUFFER_SIZE * sizeof(short));
	memcpy(rightAudioData, parm.WAVE_DATA.RightAudioBuffer, BUFFER_SIZE * sizeof(short));
	double maxAmplitude = 0;
	for (int i = 0; i < BUFFER_SIZE; ++i) {
		double sample = static_cast<double>(leftAudioData[i]);
		if (fabs(sample) > maxAmplitude) {
			maxAmplitude = fabs(sample);
		}
	}
	maxAmplitude = maxAmplitude >= 32768 ? 32767 : maxAmplitude;
	double maxAmplitude_dB = 20 * log10(maxAmplitude / 32767);

	for (int i = 0; i < BUFFER_SIZE; i++) {
		leftSpectrumData[i] = sqrt((leftSpectrum[i][0] * leftSpectrum[i][0]) + (leftSpectrum[i][1] * leftSpectrum[i][1])) / BUFFER_SIZE;;
		rightSpectrumData[i] = sqrt((rightSpectrum[i][0] * rightSpectrum[i][0]) + (rightSpectrum[i][1] * rightSpectrum[i][1])) / BUFFER_SIZE;;
	}
	//���R�W��
	ASAudio::GetInstance().SpectrumAnalysis(leftSpectrumData, rightSpectrumData);
	//�_�쭵�q
	parm.bMuteTest = false;
	if (MuteWaveIn)
		ASAudio::GetInstance().SetMicMute(false);
	if (MuteWaveOut)
		ASAudio::GetInstance().SetSpeakerSystemVolume();

	return result;
}
