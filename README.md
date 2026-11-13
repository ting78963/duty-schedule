# 交通隊勤務自動化系統

## 部署到 Render

1. 把此 repo 推上 GitHub
2. Render → New Web Service → 連結 GitHub repo
3. 設定：
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn app:app`
   - Python: 3.11
4. 環境變數：
   - `SUPABASE_URL`: 你的 Supabase URL
   - `SUPABASE_KEY`: Flask 伺服器端使用的 Supabase service_role 金鑰

## Supabase 建表

```sql
create table staff (
  code integer primary key,
  name text not null,
  type text default '輪班'
);
```

## 上傳模板

「模板管理」以管理密碼上傳已驗收的 .docx 到私有 `duty-templates` 儲存區。

| 模板 | Storage 物件名稱 |
| --- | --- |
| 平日無取酒 | `weekday.docx` |
| 假日無取酒 | `holiday.docx` |
| 平日取酒 19–23 | `weekday_sake.docx` |
| 平日取酒 18–22 | `weekday_sake_18_22.docx` |
| 假日取酒 18–22 | `holiday_sake.docx` |

新增平日 18–22 時，只上傳已驗收的完整修正版到新物件名稱。Flask 優先讀取私有 Storage，物件不存在時才使用對應的本機 `assets/` 副本。含勤務資料的本機 Word 範本由 Git 忽略，部署環境需另外備妥私有物件。
