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
APP_ENV=production
SESSION_SECRET=change-this-to-a-random-long-secret
SESSION_MAX_AGE_HOURS=12
SESSION_COOKIE_SECURE=true
```

请把 `SESSION_SECRET` 改成随机长字符串。`APP_ENV=production` 但仍使用默认密钥时，系统启动会打印明确警告。修改 `.env` 后需要重启服务才会生效。登录后后台顶部会显示当前账号，并提供“退出登录”和“切换账号”。本地 HTTP 环境下 Cookie 使用 `HttpOnly` 和 `SameSite=Lax`；正式部署到 HTTPS 后建议设置 `SESSION_COOKIE_SECURE=true`。

后台认证只读取 `admin_session` Cookie，不再接受浏览器缓存的 HTTP Basic Authorization 自动登录。点击“退出登录”或“切换账号”会删除登录 Cookie 并返回 `/admin/login?logged_out=1`，需要重新输入账号密码。

上传图片和普通文件都通过后台登录保护访问。图片地址为 `/admin/uploads/{filename}`，普通文件下载地址为 `/admin/files/{file_id}`。不要把 `uploads/` 配置成公开静态目录。

## 访问地址

- 门店提报页：http://127.0.0.1:8701/submit
- 门店查询页：http://127.0.0.1:8701/query
- 后台管理页：http://127.0.0.1:8701/admin
- 业务总览：http://127.0.0.1:8701/admin/dashboard
- Excel 导出：http://127.0.0.1:8701/admin/export（需后台登录，推荐从工单管理页右上角点击）
- 归档工单：http://127.0.0.1:8701/admin/archive
- 门店排班：http://127.0.0.1:8701/admin/schedules
- 员工管理：http://127.0.0.1:8701/admin/employees
- 班次设置：http://127.0.0.1:8701/admin/shift-types
- 门店查看排班：http://127.0.0.1:8701/schedule
- 配置管理：http://127.0.0.1:8701/admin/settings
- 嵌入页面管理：http://127.0.0.1:8701/admin/embedded-pages
- 回收站：http://127.0.0.1:8701/admin/trash
- 测试数据清理：http://127.0.0.1:8701/admin/cleanup
- 账号设置：http://127.0.0.1:8701/admin/account
- 系统设置：http://127.0.0.1:8701/admin/system
- 后台路由体检：http://127.0.0.1:8701/admin/route-health
- 运行版本诊断：http://127.0.0.1:8701/__version

Logo 文件位置：`static/img/zhiyang-logo.png`。后台侧边栏和公开页面顶部都会引用这个 Logo；如果需要更换 Logo，只替换这个文件即可。不要删除 `static/img/zhiyang-logo.png`，否则页面会回退为文字展示。

`run.bat` 启动前会打印当前目录、Git commit、Python 路径、`main.py` 路径、路由数量和关键后台路由缺失清单；如果关键路由缺失，会停止启动，避免继续运行旧版本或错误目录。

门店公开端页面统一使用顶部导航：`/submit` 可进入“查询工单、返回业务总览、返回工单管理”，`/query` 可进入“提交新工单、返回业务总览、返回工单管理”，工单详情页可返回查询结果、提交新工单或返回门店查询。

## 门店如何填写

1. 打开门店提报页。
2. 选择一个或多个门店、填写提报人。同一张工单只生成一个工单号，涉及多个门店时会同时关联这些门店。
3. 选择需求类型和紧急程度。
4. 按需选择一个或多个品牌，也可以在“其他品牌”中手动填写；商品名称、规格条码、数量、期望完成时间按实际情况填写。
5. 在问题说明中写清楚需求或异常情况。
6. 可上传多张图片，支持 `jpg`、`jpeg`、`png`、`webp`，单张不超过 10MB，默认最多 5 张、总大小不超过 30MB。
7. 可按需上传普通文件，例如 PDF、Word、Excel、CSV、TXT、ZIP、RAR。图片和文件在提交前都可以从待上传列表中删除。
8. 提交后页面会显示工单号，格式为 `REQ-YYYYMMDD-四位流水号`。

## 门店查询与进度查看

门店可以打开 `/query` 查询本店工单。门店是必填项，工单号不是必填项；只选择门店时，系统默认展示该门店最近 30 天工单，因此不强制要求门店记住工单号。涉及多个门店的同一张工单，任一关联门店都可以在本店查询结果和详情页中看到。

查询结果中可点击“查看详情”进入门店端工单详情页。详情页会展示当前状态、处理人、总部处理备注、时效状态、问题说明、附件数量和补充资料记录；门店端详情页不提供后台附件下载、删除或状态修改入口。

如果总部要求补充资料，门店可在查询结果或详情页点击“补充资料”，上传补充说明、图片或文件。补充人必填；补充说明、图片、文件至少需要提供一种。若工单状态为“待门店补充”，补充成功后会按 `config/system.json` 中的 `supplement_status_after_store_update` 自动回到默认“待处理”。门店端不能删除附件，如需删除附件或更正上传内容，请联系总部在后台处理。

## 工单协作说明

后台工单详情页提供协作层，不改变原有工单提报、处理和导出流程。总部可维护协作人、拆分子任务，并发布沟通记录；子任务支持负责人、截止时间和状态，新增或更新都会写入处理日志并生成站内消息。

沟通记录分为“门店可见”和“内部备注”。门店端详情页只展示门店可见沟通，内部备注只在后台详情页展示。门店也可以在详情页提交公开沟通内容，系统会校验工单属于当前门店后写入沟通记录、处理日志和站内消息。

## 门店排班模块

后台新增“门店排班”分组，第一版只做排班计划，不做考勤打卡、工资核算或复杂自动排班。

- 排班管理入口：`/admin/schedules`，支持单店、多店和全部门店查看；门店、员工、员工状态、班次均支持多选筛选，可在日历视图、员工视图、表格视图和门店汇总之间切换。
- 员工管理入口：`/admin/employees`，维护员工姓名、所属门店、角色、电话和状态；离职或停用员工不能继续排班，员工不会被物理删除。
- 班次设置入口：`/admin/shift-types`，班次可按门店配置，也可显式设置为通用班次；旧全局班次会自动兼容为通用班次，不同门店可以维护同名但时间不同的早班、晚班、全天或休息班次。
- 班次设置页可维护门店营业时间。未配置营业时间不会阻断排班，但页面会提示先在班次设置中维护。
- 门店查看排班入口：`/schedule`，门店公开端可选择门店和月份查看排班，只能查看，不能修改。
- 后台排班表单支持一次勾选多名员工和多天日期，为这些员工和日期批量写入同一班次；批量排班必须选择单个具体门店，多门店模式只用于查看、看板和导出。
- 排班支持自定义时间，可填写自定义名称、开始时间、结束时间和工时；工时留空时按开始/结束时间自动计算，结束时间早于开始时间时按跨天班次处理。
- 排班看板展示当月已排工时、截至当前日期已排工时、当月排班人数/人次、休息班次数、自定义/加班班次数、门店工时排行和班次分布。这里统计的是已排计划工时，不是打卡实际工时。
- 同一员工同一天只能有一个班次；旧的单人单日提交仍会更新原记录，不产生重复排班。
- 排班保存和删除会写入 `schedule_logs`，便于追踪谁在什么时候做了调整。
- 排班 Excel 导出入口：`/admin/schedules/export`，导出范围随当前页面筛选变化；单店、多店、全部门店分别使用 `门店排班_门店名_YYYY-MM.xlsx`、`门店排班_多门店_YYYY-MM.xlsx`、`门店排班_全部门店_YYYY-MM.xlsx` 文件名。

## 嵌入 HTML 更新机制

后台提供 `/admin/embedded-pages` 管理嵌入 HTML 页面。管理员可以新增页面、替换单个 HTML 或 ZIP 资源包、设置导航名称和页面标题，并启用或停用左侧导航入口。启用后的页面会显示在后台左侧“扩展页面”分组，访问路径为 `/admin/embed/{page_key}`。

嵌入页面文件保存到 `data/embedded_pages/{page_key}/`，属于运行数据，不上传 GitHub，也不会被 `git pull` 覆盖。新上传的单 HTML 会保存为 `index.html`；ZIP 会解压到同一个 page_key 目录，推荐把入口文件命名为 `index.html`，可包含 `assets/` 下的图片、CSS、JS 等资源。如果 ZIP 根目录只有一个 HTML 文件，系统会自动把它识别为入口并保存为 `index.html`；如果根目录有多个 HTML 文件，则需要手动将入口文件命名为 `index.html`。旧版本遗留的 `data/embedded_pages/{page_key}.html` 仍可兼容读取，但新上传统一使用目录结构。系统代码更新仍然需要 GitHub push、服务器 git pull 和重启服务；嵌入页面更新不需要重启服务，管理员上传或替换后，所有后台用户刷新页面即可看到新版。如果其他用户已经打开旧页面，需要刷新浏览器才会看到最新版本。iframe 地址会带版本参数，避免浏览器缓存旧 HTML；未来如果需要自动实时刷新，可以再接入消息提醒或轮询版本号。

上传限制来自 `config/system.json`：`max_embedded_html_mb` 默认 20MB，`max_embedded_zip_mb` 默认 100MB。如果确实只想上传 50MB 单 HTML，可以把 `max_embedded_html_mb` 调整到 60，但不建议长期依赖超大单 HTML；大型页面建议上传 ZIP 资源包，让 CSS、JS、图片等资源拆分到 `assets/`。ZIP 解压会拒绝路径穿越和 `exe`、`bat`、`cmd`、`sh`、`py`、`php`、`jar`、`msi` 等高风险后缀。系统不解析和清洗 HTML/JS 内容，嵌入页面被视为内部可信内容，只允许后台账号上传。不要上传不可信来源、含外部恶意脚本、账号密码、令牌或敏感数据的 HTML/ZIP；如果后续要开放给非技术人员使用，需要增加更严格的 HTML 清洗和审核。

## 附件上传说明

图片上传和文件上传是两个入口：

- 图片用于问题截图、现场照片等，仍走 `images` 字段。
- 普通文件用于表格、文档、压缩包等，走 `files` 字段。
- 图片和文件在提交前都可以从待上传列表中删除。
- 提交成功后，门店不能自行删除附件。如发现提交内容或附件错误，需要联系总部在后台删除，不要重复提交多次。
- 后台查看图片、下载文件、删除附件都需要登录。
- 普通文件允许格式和大小限制来自 `config/system.json`：`allowed_file_extensions`、`max_file_mb`、`max_file_count`、`max_total_file_upload_mb`。
- `uploads/` 不要上传 GitHub。
- `backup.sh` 会备份 `uploads/` 和 `data/embedded_pages/`，因此普通文件附件和嵌入 HTML 运行数据都会被一起备份。

## 总部如何处理

1. 打开后台管理页。后台采用“止痒 ERP”布局，左侧导航包含业务总览、工单管理、归档工单、门店查询、配置管理、账号设置和系统设置。数据导出入口统一放在工单管理页右上角。
2. 可按门店、需求类型、紧急程度、状态、日期范围、关键词筛选；多门店工单会出现在每个关联门店的筛选结果中。
3. 点击工单号进入详情页。
4. 在详情页选择处理人、修改状态和处理备注。
5. 点击保存后，系统会记录最后更新时间；状态改为“已完成”时会自动写入完成时间，重新打开时会清空完成时间。
6. 每次状态、处理人或备注发生变化，详情页底部会保留处理日志。
7. 详情页可查看图片和文件附件，普通文件通过 `/admin/files/{file_id}` 下载。
8. 总部可在详情页补充上传处理凭证、截图、表格或供应商反馈文件，补充上传会写入处理日志。
9. 如门店提交错附件，总部可在详情页删除图片或文件，删除前会有确认提示，删除操作会写入处理日志。
10. 后台列表和详情页会显示时效状态：未设置、已超时、今日到期、未到期、超时完成、按时完成。列表支持按时效筛选。
11. 后台 `/admin/dashboard` 提供业务总览看板，可按日期、门店、需求类型、状态和处理人筛选，并展示核心指标、最近工单动态和工单类型结构；门店统计会把多门店工单分别计入每个关联门店。
12. 工单管理页支持多选、当前页全选、按当前筛选条件选择全部结果，并可批量归档或批量移入回收站。批量操作会要求确认并写入工单日志。
13. `/admin/archive` 是归档页，用于查看、筛选、导出、恢复归档工单，或把归档工单移入回收站。
14. `/admin/trash` 是回收站，用于恢复已删除工单和嵌入页面，或在确认后永久删除。
15. `/admin/cleanup` 是测试数据清理页，必须先预览匹配工单，再确认批量移入回收站。
16. `/admin/settings`、`/admin/account`、`/admin/system` 是登录后可见的占位页，用于查看本地配置、账号和运行维护提示。

## 归档与回收站

工单分为三种数据状态：

- 活跃工单：`deleted_at` 为空且 `archived_at` 为空，默认显示在 `/admin` 工单管理页，用于日常处理。
- 归档工单：`deleted_at` 为空且 `archived_at` 不为空，不显示在默认活跃列表，显示在 `/admin/archive`。归档适合已完成、已驳回、历史沉淀、暂不处理但需要长期保留的工单。
- 回收站工单：`deleted_at` 不为空，不显示在活跃列表，也不显示在门店端 `/query`，显示在 `/admin/trash`。回收站适合误提、测试、垃圾工单。

归档不是删除。归档数据会长期保留，可查询、可导出、可恢复到活跃列表；门店端仍可查询归档工单，页面会显示“已归档”标记。回收站工单不会出现在门店端查询、默认后台统计和活跃导出中。

工单管理页支持多选、当前页全选、清空选择、按当前筛选条件选择全部结果。选择全部筛选结果时，后端会按当前筛选条件执行批量归档或批量移入回收站，不只处理当前页。所有批量操作都需要后台登录、CSRF、确认提示，并写入 `ticket_logs`。

归档页支持筛选、分页、多选、批量恢复到活跃列表、批量移入回收站，以及导出当前筛选范围的归档工单。回收站支持筛选、分页、多选、批量恢复和批量永久删除；如果归档工单先移入回收站，恢复后仍回到归档页，活跃工单恢复后回到活跃列表。

永久删除不可恢复，会真正删除数据库记录、关联附件记录和对应物理文件。执行永久删除前建议先运行 `backup.sh` 备份 `data/tickets.db`、`uploads/` 和 `data/embedded_pages/`。

Excel 导出范围随当前页面和筛选条件变化：`/admin/export` 导出活跃工单当前筛选结果，`/admin/archive/export` 导出归档工单当前筛选结果。回收站默认不提供常规 Excel 导出。

后台详情页的协作人、评论、子任务和门店补充记录支持软删除。删除后不再出现在后台协作区或门店详情页；相关删除动作仍会写入处理日志，但日志展示不会继续暴露已隐藏评论、子任务、协作人或补充说明的原文。

嵌入页面删除也先进入回收站。软删除后页面会自动停用，左侧扩展导航和 `/admin/embed/{page_key}` 不再显示；恢复后会重新启用。永久删除嵌入页面会移除数据库记录和 `data/embedded_pages/{page_key}/` 资源目录。

测试数据清理页适合清理 Codex、自测、smoke 等临时工单。建议保留“只匹配测试数据”勾选，先预览结果，再批量移入回收站。批量清理不会直接永久删除，正式数据如需清理应先备份并人工核对筛选条件。

状态包括：

- 待处理
- 处理中
- 待门店补充
- 已完成
- 已驳回

## 消息提醒说明

后台页面提供轻量级站内消息提醒，不接短信、微信、飞书，也不使用 WebSocket。后台页面打开后会每 15 秒检查一次新消息；首次打开只刷新未读数量和消息列表，不会把历史消息全部弹出来。

以下事件会生成后台消息：

- 门店提交新工单。
- 门店补充资料。
- 工单状态从其他状态变为“待门店补充”。

后台右上角“消息”按钮会显示未读数量，点击后可查看最近消息、进入工单详情、标记单条已读或全部已读。有新消息到达时，页面右下角会出现站内弹窗；“当天必须处理”的新工单会以更醒目的紧急样式显示。

浏览器桌面通知需要后台人员主动点击“开启桌面提醒”并授权。授权后，新消息到达时会同时触发浏览器通知；通知点击后会打开对应工单。桌面通知需要后台页面保持打开，不是手机微信式离线推送。如果后续需要外部提醒，可以再接企业微信或飞书 webhook。

## 如何导出 Excel

在后台工单管理页右上角点击“导出 Excel”，或登录后直接访问：

```text
http://127.0.0.1:8701/admin/export
```

从工单管理页点击导出时，会自动携带当前门店、状态、时效、日期、关键词等筛选条件；导出的是当前活跃工单筛选结果的全部工单，不受当前分页影响。归档页右上角的“导出归档 Excel”会访问 `/admin/archive/export`，只导出归档页当前筛选范围。左侧导航不再单独显示“数据导出”菜单。

多门店、多品牌工单导出时仍保持一行一个工单，门店和品牌字段会以合并文本展示，方便总部继续用 Excel 汇总筛选。

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
config/request_type_rules.json
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

### 如何维护需求类型必填规则

编辑 `config/request_type_rules.json`，可为不同需求类型配置额外必填项、图片/文件要求和说明提示：

```json
{
  "建单需求": {
    "required_fields": ["brand", "product_name", "quantity"],
    "require_image": false,
    "require_file": false,
    "require_any_attachment": true,
    "description_hint": "请说明到货情况、采购单需求或建单原因"
  }
}
```

`required_fields` 支持 `brand`、`product_name`、`sku_barcode`、`quantity`、`description`、`expected_finish_date`。如果规则文件不存在、JSON 格式错误，或某个需求类型没有配置规则，系统只使用基础校验，不会阻断提交。

### 如何新增品牌

编辑 `config/brands.json`，按 JSON 数组格式增加品牌名称，例如：

```json
[
  "自有品牌",
  "联名品牌"
]
```

门店提报页会把这里的品牌显示为可多选选项；未在列表中的品牌仍可在“其他品牌”中手动填写。系统会继续把品牌合并保存到原有工单字段，同时写入多品牌关联表，便于详情、查询和导出统一展示。

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

处理人列表会自动合并 `config/handlers.json` 与 `.env` 中的 `ADMIN_USERS` 用户名；如果 `config/handlers.json` 不存在，则直接使用 `ADMIN_USERS` 作为处理人列表。这样能登录后台的账号默认都可以被选择为处理人。旧的处理人名称（例如“总部运营”）继续通过 `config/handlers.json` 保留。

后台也提供登录态接口 `GET /api/handlers`，返回当前统一处理人列表，便于前端控件使用同一来源。

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
  "max_total_file_upload_mb": 50,
  "max_embedded_html_mb": 20,
  "max_embedded_zip_mb": 100,
  "store_query_default_days": 30,
  "store_query_page_size": 20,
  "supplement_status_after_store_update": "待处理",
  "max_bulk_schedule_count": 200
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
- `max_embedded_html_mb`：后台嵌入 HTML 单文件大小限制，默认 20MB。
- `max_embedded_zip_mb`：后台嵌入 ZIP 资源包大小限制，默认 100MB。推荐 ZIP 根目录入口文件命名为 `index.html`；如果根目录只有一个 HTML 文件，系统会自动识别并保存为 `index.html`。ZIP 可包含 `assets/` 图片、CSS、JS 等资源。
- `store_query_default_days`：门店只选择门店查询时默认回看天数，默认 30。
- `store_query_page_size`：门店查询页每页条数，默认 20。
- `supplement_status_after_store_update`：门店补充“待门店补充”工单后自动切回的状态，默认“待处理”。

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
- 嵌入 HTML 运行数据：`data/embedded_pages/`

不要随意删除数据库文件、上传附件目录或嵌入 HTML 目录。删除数据库会清空所有工单记录，删除 `uploads/` 会导致历史工单图片无法查看、普通文件无法下载，删除 `data/embedded_pages/` 会导致后台扩展页面无法显示。

## 数据库升级说明

系统启动时会自动检查并补齐：

- `tickets.assigned_to`
- `tickets.closed_at`
- `ticket_logs`
- `ticket_files`
- `ticket_supplements`
- `ticket_images.source`
- `ticket_images.uploaded_by`
- `ticket_images.supplement_id`
- `ticket_files.source`
- `ticket_files.uploaded_by`
- `ticket_files.supplement_id`
- `ticket_participants`
- `ticket_comments`
- `ticket_tasks`
- `embedded_pages`
- `tickets.deleted_at`、`tickets.deleted_by`、`tickets.delete_reason`
- `tickets.archived_at`、`tickets.archived_by`、`tickets.archive_reason`
- `ticket_supplements.deleted_at`、`ticket_supplements.deleted_by`、`ticket_supplements.delete_reason`
- `ticket_participants.deleted_at`、`ticket_participants.deleted_by`、`ticket_participants.delete_reason`
- `ticket_comments.deleted_at`、`ticket_comments.deleted_by`、`ticket_comments.delete_reason`
- `ticket_tasks.deleted_at`、`ticket_tasks.deleted_by`、`ticket_tasks.delete_reason`
- `embedded_pages.deleted_at`、`embedded_pages.deleted_by`、`embedded_pages.delete_reason`

升级过程是兼容式迁移，不会清空 `tickets`、`ticket_images`，也不会删除 `uploads/` 或 `data/embedded_pages/`。上线前仍建议先备份数据库、附件和嵌入 HTML 运行数据。

## Nginx + HTTPS 正式部署

测试阶段可以临时访问 `http://公网IP:8701`。正式给门店长期使用前，建议配置域名、Nginx 和 HTTPS，并让 uvicorn 只监听 `127.0.0.1:8701`，不要长期把 8701 端口直接暴露到公网。

项目提供示例文件：

```text
deploy/nginx-store-request-tool.conf.example
```

示例包含 80 跳转 HTTPS、443 ssl、`proxy_pass http://127.0.0.1:8701`、`client_max_body_size 120M` 和常用安全头。云端上传 ZIP 或大 HTML 时，Nginx 的 `client_max_body_size` 必须大于实际上传文件大小；例如默认 100MB ZIP，建议配置 120M 或 150M。证书路径和 `server_name request.example.com` 需要替换成自己的域名，可使用 Let's Encrypt / certbot 申请证书。国内服务器绑定域名通常需要先完成备案。

## 备份说明

项目提供 `backup.sh`，用于备份 `data/tickets.db`、`uploads/` 和 `data/embedded_pages/`。正式使用后建议每天定时执行一次，例如通过 crontab 调度。图片和普通文件都保存在 `uploads/`，嵌入 HTML 页面保存在 `data/embedded_pages/`，因此都会进入备份。
