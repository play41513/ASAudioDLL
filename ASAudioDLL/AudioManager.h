#pragma once

#include "ConfigManager.h" // �ݭn�Ψ� Config ���c
#include <string>

/*
 * AudioManager ���O
 * ¾�d�G�@�����լy�{����ժ̡C
 * ���ھڶǤJ���]�w (Config)�A�̧ǰ���U�����T���աA
 * �ú޲z���չL�{�������A�P�̲׵��G�C
 */
class AudioManager {
public:
    // �غc�禡
    AudioManager();

    // �D�n������禡
    // �ھڴ��Ѫ��]�w�ɤ��e�A����Ҧ��Q�ҥΪ����աC
    // �p�G�Ҧ��ҥΪ����ճ����Q����A�h��^ true�C
    // �`�N�G��^ true ���N����� "�q�L (PASS)"�A�ȥN��y�{ "���槹��"�C
    // ���ժ����鵲�G (PASS/FAIL) �x�s�b resultString ���C
    bool ExecuteTestsFromConfig(const Config& config);

    // ������յ�����A�̲ײ��ͪ����G�r�� (�Ҧp "LOG:PASS#" �ο��~�T��)
    const std::string& GetResultString() const;

private:
    // �w��C�دS�w���ժ��p�����U�禡�A�� ExecuteTestsFromConfig �I�s

    // ���� THD+N�BdB�B�W�v���֤߭��T���R����
    bool RunAudioTest(const Config& config);
    // ���� WAV �ɮת�����P����
    bool RunWavPlayback(const Config& config);
    // ���歵�T�j�� (Loopback) ����
    bool RunAudioLoopback(const Config& config);
    // ��������t�ιw�]���T�˸m���\��
    bool RunSwitchDefaultDevice(const Config& config);

    // �b�Ҧ����ն}�l�e�A�w���ˬd���w�����T��X/��J�˸m�O�_�s�b
    bool CheckAudioDevices(const Config& config);

    // �Ω��x�s��Ӵ��լy�{�̲׵��G���r��
    std::string resultString;
};