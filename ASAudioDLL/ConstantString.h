#pragma once
#include <unordered_map>

#define ERROR_CODE_NOT_FIND_WAVEIN_DEV			1
#define ERROR_CODE_OPEN_WAVEIN_DEV				2
#define ERROR_CODE_NOT_FIND_WAVEOUT_DEV			3
#define ERROR_CODE_OPEN_INIT_WAVEOUT_DEV		4
#define ERROR_CODE_OPEN_WAVEOUT_DEV				5

#define ERROR_CODE_START_WAVEOUT_DEV			6
#define ERROR_CODE_SET_SPEAKER_SYSTEM_VOLUME	7
#define ERROR_CODE_UNSET_PARAMETER				8

static const int BUFFER_SIZE = 2205; //������W�ШC�@���I��20Hz�A�C���Q���I��1K Hz

#define M_PI 3.14159235358979323
#define SAMPLE_RATE 44100       // �w�q�ļ˲v��44100Hz
#define AMPLITUDE 0.5           // �w�q�H�����T��0.5
#define FRAMES_PER_BUFFER 256   // �w�q�C�ӽw�İϪ��V��

// �w�q���~�X�T�|���A
enum class ErrorCode {
    NOT_FIND_WAVEIN_DEV = 1,
    OPEN_WAVEIN_DEV,
    NOT_FIND_WAVEOUT_DEV,
    OPEN_INIT_WAVEOUT_DEV,
    OPEN_WAVEOUT_DEV,
    START_WAVEOUT_DEV,
    SET_SPEAKER_SYSTEM_VOLUME,
    UNSET_PARAMETER
};

// �w�q���~�X�M���~�T�����M�g
extern const std::unordered_map<ErrorCode, const wchar_t*> error_map;
