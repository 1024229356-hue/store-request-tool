# 门店需求工单系统

这是一个轻量级本地 Web 工单工具，用于门店向总部提交工作需求，总部在后台查看、处理、跟踪状态，并导出 Excel。

## 如何启动

方式一：双击 `run.bat`。

方式二：在项目目录执行：

```bat
python -m uvicorn main:app --host 127.0.0.1 --port 8701
```

首次启动会自动初始化 `data/tickets.db`，并自动创建 `uploads` 目录。
后续版本升级时，系统启动也会自动补齐缺失的数据表和字段，不需要删除数据库。

生产环境启动方式：

```bash
python -m uvicorn main:app --host 0.0.0.0 --port 8701
```

后台访问需要登录。系统使用 Cookie 登录会话，账号密码从环境变量读取：

```text
ADMIN_USERNAME
ADMIN_PASSWORD
```

也可以使用多账号配置：

```text
ADMIN_USERS=admin:123456,caigou:123456,yunying:123456
```

生产环境必须配置 Cookie 签名密钥：

```text
SESSION_SECRET=change-this-to-a-random-long-secret
```

请改成随机长字符串。修改 `.env` 后需要重启服务才会生效。登录后后台顶部会显示当前账号，并提供“退出登录”和“切换账号”。本地 HTTP 环境下 Cookie 使用 `HttpOnly` 和 `SameSite=Lax`；正式部署到 HTTPS 后建议在反向代理和应用配置中启用 Secure Cookie。

上传图片和普通文件都通过后台登录保护访问。图片地址为 `/admin/uploads/{filename}`，普通文件下载地址为 `/admin/files/{file_id}`。不要把 `uploads/` 配置成公开静态目录。

## 访问地址

- 门店提报页：http://127.0.0.1:8701/submit
- 后台管理页：http://127.0.0.1:8701/admin
- Excel 导出：http://127.0.0.1:8701/admin/export

## 门店如何填写

1. 打开门店提报页。
2. 选择门店、填写提报人。
3. 选择需求类型和紧急程度。
4. 按需填写品牌、商品名称、规格条码、数量、期望完成时间。
5. 在问题说明中写清楚需求或异常情况。
6. 可上传多张图片，支持 `jpg`、`jpeg`、`png`、`webp`，单张不超过 10MB，默认最多 5 张、总大小不超过 30MB。
7. 可按需上传普通文件，例如 PDF、Word、Excel、CSV、TXT、ZIP、RAR。图片和文件在提交前都可以从待上传列表中删除。
8. 提交后页面会显示工单号，格式为 `REQ-YYYYMMDD-四位流水号`。

## 附件上传说明

图片上传和文件上传是两个入口：

- 图片用于问题截图、现场照片等，仍走 `images` 字段。
- 普通文件用于表格、文档、压缩包等，走 `files` 字段。
- 图片和文件在提交前都可以从待上传列表中删除。
- 提交成功后，门店不能自行删除附件。如发现提交内容或附件错误，需要联系总部在后台删除，不要重复提交多次。
- 后台查看图片、下载文件、删除附件都需要登录。
- 普通文件允许格式和大小限制来自 `config/system.json`：`allowed_file_extensions`、`max_file_mb`、`max_file_count`、`max_total_file_upload_mb`。
- `uploads/` 不要上传 GitHub。
- `backup.sh` 会备份 `uploads/`，因此普通文件附件也会被一起备份。

## 总部如何处理

1. 打开后台管理页。
2. 可按门店、需求类型、紧急程度、状态、日期范围、关键词筛选。
3. 点击工单号进入详情页。
4. 在详情页选择处理人、修改状态和处理备注。
5. 点击保存后，系统会记录最后更新时间；状态改为“已完成”时会自动写入完成时间，重新打开时会清空完成时间。
6. 每次状态、处理人或备注发生变化，详情页底部会保留处理日志。
7. 详情页可查看图片和文件附件，普通文件通过 `/admin/files/{file_id}` 下载。
8. 总部可在详情页补充上传处理凭证、截图、表格或供应商反馈文件，补充上传会写入处理日志。
9. 如门店提交错附件，总部可在详情页删除图片或文件，删除前会有确认提示，删除操作会写入处理日志。

状态包括：

- 待处理
- 处理中
- 待门店补充
- 已完成
- 已驳回

## 如何导出 Excel

在后台管理页点击“导出 Excel”，或直接访问：

```text
http://127.0.0.1:8701/admin/export
```

导出文件名格式：

```text
门店需求工单_YYYYMMDD_HHMM.xlsx
```

如果当前没有工单，也会导出只有表头的 Excel 文件。

## 阿里云 Ubuntu 部署

以下示例以 Ubuntu 24.04 和项目目录 `/opt/store-request-tool` 为例。若使用其他目录，请同步修改 systemd 服务文件中的 `WorkingDirectory`、`EnvironmentFile` 和 `ExecStart`。

1. 安装基础环境并拉取代码：

```bash
sudo apt update
sudo apt install -y git python3 python3-venv
cd /opt
sudo git clone https://github.com/1024229356-hue/store-request-tool.git
sudo chown -R $USER:$USER /opt/store-request-tool
cd /opt/store-request-tool
```

2. 创建 `.env`：

```bash
cp .env.example .env
nano .env
```

将示例密码改成自己的后台密码：

```text
ADMIN_USERNAME=admin
ADMIN_PASSWORD=change-me
SESSION_SECRET=change-this-to-a-random-long-secret
```

正式使用建议改为多账号：

```text
ADMIN_USERS=admin:123456,caigou:123456,yunying:123456
SESSION_SECRET=change-this-to-a-random-long-secret
```

`SESSION_SECRET` 必须改成随机长字符串。修改 `.env` 后执行 `sudo systemctl restart store-request-tool` 重启服务。

3. 运行部署脚本：

```bash
chmod +x deploy.sh backup.sh
./deploy.sh
sudo chown -R www-data:www-data /opt/store-request-tool
```

4. 配置 systemd：

```bash
sudo cp store-request-tool.service.example /etc/systemd/system/store-request-tool.service
sudo nano /etc/systemd/system/store-request-tool.service
sudo systemctl daemon-reload
sudo systemctl enable store-request-tool
```

确认服务文件中的项目目录和运行用户正确后启动：

```bash
sudo systemctl start store-request-tool
```

5. 常用运维命令：

```bash
sudo systemctl stop store-request-tool
sudo systemctl restart store-request-tool
sudo systemctl status store-request-tool
sudo journalctl -u store-request-tool -f
```

6. 阿里云防火墙放行 `8701`：

- 在阿里云 ECS 安全组入方向放行 TCP `8701`。
- 如果服务器启用了 UFW，也执行：

```bash
sudo ufw allow 8701/tcp
```

完成后访问：

```text
http://服务器公网IP:8701/submit
http://服务器公网IP:8701/admin
```

访问后台时会先进入 `/admin/login` 登录页。登录后可在后台顶部退出登录或切换账号；Excel 导出、图片查看和文件下载也会校验同一个后台登录 Cookie。

### 阿里云更新部署命令

已有线上目录时，可按下面命令更新：

```bash
cd /opt/store-request-tool
sudo git pull
sudo ./deploy.sh
sudo chown -R www-data:www-data /opt/store-request-tool
sudo systemctl restart store-request-tool
sudo systemctl status store-request-tool
```

不建议长期直接使用 `http://公网IP:8701/admin` 暴露后台。正式试运行后，建议配置 Nginx + HTTPS，并限制后台访问入口；HTTPS 后建议启用 Secure Cookie。

## 配置文件说明

业务配置统一放在 `config/` 目录：

```text
config/stores.json
config/request_types.json
config/urgency_levels.json
config/statuses.json
config/brands.json
config/handlers.json
config/system.json
```

如果配置文件不存在、JSON 格式错误或核心配置为空，系统会自动使用内置默认值兜底，不会因为配置问题启动失败。

### 如何新增门店

编辑 `config/stores.json`，按 JSON 数组格式增加门店名称，例如：

```json
[
  "南京门东店",
  "南昌万寿宫店",
  "新门店名称"
]
```

### 如何新增需求类型

编辑 `config/request_types.json`，按 JSON 数组格式增加需求类型，例如：

```json
[
  "建单需求",
  "审单需求",
  "门店陈列需求"
]
```

新增后，门店提报页和后台筛选会使用新的需求类型；提交校验也会按这个文件判断。

### 如何新增品牌

编辑 `config/brands.json`，按 JSON 数组格式增加品牌名称，例如：

```json
[
  "自有品牌",
  "联名品牌"
]
```

品牌仍然保存到原有工单字段，页面会把这里的品牌作为输入建议。

### 如何维护处理人

编辑 `config/handlers.json`，按 JSON 数组格式维护总部处理人或处理小组：

```json
[
  "总部商品",
  "总部运营",
  "采购",
  "财务"
]
```

如果配置为空，系统仍可使用，详情页的处理人可以不指定。

### 后台多账号配置

后台账号配置在 `.env` 中维护，`.env` 不要上传 GitHub。正式试运行建议使用 `ADMIN_USERS`：

```text
ADMIN_USERS=admin:123456,caigou:123456,yunying:123456
```

- 多账号之间用英文逗号分隔。
- 用户名和密码之间用英文冒号分隔。
- 用户名和密码前后空格会自动去掉。
- 修改 `.env` 后需要重启 `run.bat` 或 systemd 服务。
- 本版本多个账号权限相同，不区分角色。

旧的单账号配置仍然兼容：

```text
ADMIN_USERNAME=admin
ADMIN_PASSWORD=change-me
```

### 如何修改系统配置

编辑 `config/system.json`：

```json
{
  "max_image_mb": 10,
  "max_image_count": 5,
  "max_total_upload_mb": 30,
  "page_size": 50,
  "allowed_image_extensions": ["jpg", "jpeg", "png", "webp"],
  "allowed_file_extensions": ["pdf", "doc", "docx", "xls", "xlsx", "csv", "txt", "zip", "rar"],
  "max_file_mb": 20,
  "max_file_count": 5,
  "max_total_file_upload_mb": 50
}
```

- `max_image_mb`：单张图片大小限制，默认 10MB。
- `max_image_count`：每张工单最多上传图片数量，默认 5。
- `max_total_upload_mb`：每张工单图片总大小限制，默认 30MB。
- `page_size`：后台列表每页工单数，默认 50。
- `allowed_image_extensions`：允许上传的图片后缀，默认 `jpg`、`jpeg`、`png`、`webp`。
- `allowed_file_extensions`：允许上传的普通文件后缀，默认 `pdf`、`doc`、`docx`、`xls`、`xlsx`、`csv`、`txt`、`zip`、`rar`。
- `max_file_mb`：单个普通文件大小限制，默认 20MB。
- `max_file_count`：每张工单最多上传普通文件数量，默认 5。
- `max_total_file_upload_mb`：每张工单普通文件总大小限制，默认 50MB。

普通文件格式必须走白名单。即使误配到白名单里，系统也会拒绝 `exe`、`bat`、`cmd`、`js`、`py`、`sh`、`php`、`jar`、`msi` 等可执行或脚本类后缀。

服务端会使用 Pillow 校验图片是否真实可打开，并检查后缀与图片格式是否匹配。

### 如何修改 Excel 文件名前缀

编辑 `config/system.json` 中的 `excel_filename_prefix`：

```json
{
  "excel_filename_prefix": "门店需求工单"
}
```

导出文件名格式为：

```text
配置的前缀_YYYYMMDD_HHMM.xlsx
```

### 哪些配置改完需要重启服务

- 通常不需要重启：`stores.json`、`request_types.json`、`urgency_levels.json`、`statuses.json`、`brands.json`、`handlers.json`、`system.json` 中的图片/文件上传限制、分页大小、默认状态、导出文件名前缀。
- 建议重启后生效：`system.json` 中的 `app_name`，以及 systemd 服务文件、`.env` 后台账号密码、端口监听方式等启动级配置。

## 数据和附件位置

- 数据库：`data/tickets.db`
- 上传图片和普通文件：`uploads/`

不要随意删除数据库文件或上传附件目录。删除数据库会清空所有工单记录，删除 `uploads/` 会导致历史工单图片无法查看、普通文件无法下载。

## 数据库升级说明

系统启动时会自动检查并补齐：

- `tickets.assigned_to`
- `tickets.closed_at`
- `ticket_logs`
- `ticket_files`

升级过程是兼容式迁移，不会清空 `tickets`、`ticket_images`，也不会删除 `uploads/`。上线前仍建议先备份数据库和附件。

## 备份说明

项目提供 `backup.sh`，用于备份 `data/tickets.db` 和 `uploads/`。正式使用后建议每天定时执行一次，例如通过 crontab 调度。图片和普通文件都保存在 `uploads/`，因此都会进入备份。
