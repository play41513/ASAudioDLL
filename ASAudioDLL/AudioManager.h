#pragma once

#include "ConfigManager.h" // 需要用到 Config 結構
#include <string>

/*
 * AudioManager 類別
 * 職責：作為測試流程的協調者。
 * 它根據傳入的設定 (Config)，依序執行各項音訊測試，
 * 並管理測試過程中的狀態與最終結果。
 */
class AudioManager {
public:
    // 建構函式
    AudioManager();

    // 主要的執行函式
    // 根據提供的設定檔內容，執行所有被啟用的測試。
    // 如果所有啟用的測試都順利執行，則返回 true。
    // 注意：返回 true 不代表測試 "通過 (PASS)"，僅代表流程 "執行完畢"。
    // 測試的具體結果 (PASS/FAIL) 儲存在 resultString 中。
    bool ExecuteTestsFromConfig(const Config& config);

    // 獲取測試結束後，最終產生的結果字串 (例如 "LOG:PASS#" 或錯誤訊息)
    const std::string& GetResultString() const;

private:
    // 針對每種特定測試的私有輔助函式，由 ExecuteTestsFromConfig 呼叫

    // 執行 THD+N、dB、頻率等核心音訊分析測試
    bool RunAudioTest(const Config& config);
    // 執行 WAV 檔案的播放與停止
    bool RunWavPlayback(const Config& config);
    // 執行音訊迴路 (Loopback) 測試
    bool RunAudioLoopback(const Config& config);
    // 執行切換系統預設音訊裝置的功能
    bool RunSwitchDefaultDevice(const Config& config);

    // 在所有測試開始前，預先檢查指定的音訊輸出/輸入裝置是否存在
    bool CheckAudioDevices(const Config& config);

    // 用於儲存整個測試流程最終結果的字串
    std::string resultString;
};