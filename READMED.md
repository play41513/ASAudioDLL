\# ASAudioTest - Factory Audio Analysis Module



ASAudioTest 專為工廠量產線 (Mass Production) 設計的 C++ 音訊硬體自動化測試函式庫 (DLL)。本模組透過 Software Loopback 機制與快速傅立葉轉換 (FFT)，取代昂貴的傳統音訊分析儀，快速篩選出麥克風與揚聲器的硬體瑕疵。



\## Core Features (核心功能)



\* \*\*高精度頻譜分析\*\*：整合 `FFTW3` 函式庫，並採用 \*\*Blackman-Harris 4-term 視窗函數\*\*，有效抑制頻譜洩漏 (Spectral Leakage)，大幅提升 THD+N 在軟體測試上的精準度。

\* \*\*低頻底噪濾除\*\*：內建 20Hz 高通濾波 (High-pass filter) 邏輯，自動過濾產線常見的環境低頻干擾與 USB DC 偏壓漂移。

\* \*\*產線專屬診斷引擎 (Auto-Diagnosis)\*\*：自動將 THD+N、dBFS 破音、頻率偏移等工程數據，轉換為「白話文排障指示」(如：音量過大、左右聲道接反、接觸不良)，直接顯示於 UI 上，降低作業員誤判率。

\* \*\*C-Style API 介面\*\*：提供極度扁平化的 API，完美相容於 C# (.NET / WinForms / MAUI)、Python 或各類產線自動化控制軟體。



\## Core Metrics (測試指標)



模組單次呼叫即可同步計算以下四大關鍵指標：

1\. \*\*Frequency (頻率)\*\*：驗證取樣率精準度，防呆左右聲道接反。

2\. \*\*Fundamental Level (基頻電平 / dBFS)\*\*：驗證麥克風/揚聲器增益，檢測削峰失真 (Clipping)。

3\. \*\*THD+N (總諧波失真加雜訊)\*\*：評估音質純淨度，抓出空焊、OP-AMP 損壞或單體破音。

4\. \*\*SNR (訊噪比)\*\*：檢測環境電流聲、接地干擾 (Ground Loop) 或漏電問題。



\## 🛠️ Tech Stack (技術堆疊)



\* \*\*C++14 / 17\*\*

\* \*\*Windows API\*\*: 

&#x20; \* `WaveIn` / `WaveOut` (Multimedia API)

&#x20; \* `Core Audio API` (Endpoint Volume / IMMDevice)

&#x20; \* `DirectShow` (Audio Loopback)

\* \*\*Audio Engine\*\*: `PortAudio` (精準正弦波生成)

\* \*\*Math \& DSP\*\*: `FFTW3` (Fast Fourier Transform)

