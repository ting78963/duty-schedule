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
   - `SUPABASE_KEY`: 你的 Supabase anon key

## Supabase 建表

```sql
create table staff (
  code integer primary key,
  name text not null,
  type text default '輪班'
);
```

## 上傳模板

部署後進入「模板管理」頁面，上傳四個 .docx 模板：
- 平日
- 假日
- 平日取締酒駕
- 假日取締酒駕
