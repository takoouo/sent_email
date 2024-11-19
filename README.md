# sent_email
## 內含檔案：
  main.py
  
  sent.py
  
  111數位精進方案行動載具狀態.xlsx // 對Serial Number跟人名的檔案 => name_file
  
  20241118_載具清單_353301.xlsx // 檢查時間的檔案 => time_file
  
  contacts.csv // 通訊錄
  
## 使用方式
  1. 檢查是否包含以下欄位
     
    name_file => 領用人、備註

    time_file => 最後連線
  
  2. 開啟main.py，並選擇使用的檔案
     
     => 選擇檢查年分 + 月份，以空格分開 (如2024 11)
     
     => 輸出結果將寫入output_mail.json
     
  3. 開啟sent.py
     
     => 檢查py內的sender_name、jsonFile路徑及檔名、標題subject及email_template.txt內文的變數名稱是否與variables相同
     
     => 執行
     
