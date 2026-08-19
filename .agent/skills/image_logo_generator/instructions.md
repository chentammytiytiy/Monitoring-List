# Image Logo Generator Skill

## Description
此 Skill 的主要用途是：當你收到生成圖片的指令時，確保生成的圖片中會帶有指定的 Logo。

## Instructions
當你需要執行圖片生成任務時，請嚴格遵守以下步驟：
1. **使用正確的工具**：一律使用 `generate_image` 工具來生成圖片。
2. **設定提示詞 (Prompt)**：根據使用者的需求撰寫提示詞，並在提示詞中明確加上「請在圖片中自然地融入或帶入提供的 Logo 圖案」之類的描述。
3. **附加圖片參數 (ImagePaths)**：在調用 `generate_image` 工具時，必須在 `ImagePaths` 參數中傳入此 Logo 的絕對路徑。
4. **Logo 檔案路徑**：指定的 Logo 圖片存放於以下絕對路徑：
   `d:\AI\AI_CLASS\.agent\skills\image_logo_generator\logo.png`

## Exception Handling
在生成圖片前，若系統無法讀取到上述的 `logo.png`，請提醒使用者將他們想要使用的 Logo 圖片命名為 `logo.png` 並放置於 `d:\AI\AI_CLASS\.agent\skills\image_logo_generator/` 目錄下。
