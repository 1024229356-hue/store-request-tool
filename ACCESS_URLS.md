# 止痒 ERP 访问地址

## 本地服务

- 门店提交：http://127.0.0.1:8701/submit
- 门店查询：http://127.0.0.1:8701/query
- 门店查看排班：http://127.0.0.1:8701/schedule
- 后台登录：http://127.0.0.1:8701/admin/login
- 业务总览：http://127.0.0.1:8701/admin/dashboard
- 工单管理：http://127.0.0.1:8701/admin
- 我的待办：http://127.0.0.1:8701/admin/my-work
- 归档工单：http://127.0.0.1:8701/admin/archive
- 回收站：http://127.0.0.1:8701/admin/trash
- 测试数据清理：http://127.0.0.1:8701/admin/cleanup
- 员工管理：http://127.0.0.1:8701/admin/employees
- 班次设置：http://127.0.0.1:8701/admin/shift-types
- 门店排班：http://127.0.0.1:8701/admin/schedules
- 配置管理：http://127.0.0.1:8701/admin/settings
- 账号设置：http://127.0.0.1:8701/admin/account
- 系统设置：http://127.0.0.1:8701/admin/system
- 嵌入页面管理：http://127.0.0.1:8701/admin/embedded-pages
- 路由体检：http://127.0.0.1:8701/admin/route-health
- 当前版本：http://127.0.0.1:8701/__version
- 健康检查：http://127.0.0.1:8701/healthz

## 阿里云部署后

- 使用 http://服务器IP:8701
- 或 https://你的域名
- 路径保持一致，只替换域名或服务器 IP。

## 兼容地址

以下地址仅用于兼容旧链接，不建议长期收藏。系统会返回 303 并跳转到标准入口。

- /admin/personnel -> /admin/employees
- /admin/staff -> /admin/employees
- /admin/employee -> /admin/employees
- /admin/schedule -> /admin/schedules
- /admin/store-schedule -> /admin/schedules
- /admin/shift-type -> /admin/shift-types
- /admin/archive-list -> /admin/archive
- /admin/recycle -> /admin/trash
- /admin/trashes -> /admin/trash
- /admin/home -> /admin/dashboard
- /admin/index -> /admin/dashboard
- /admin/tickets -> /admin
- /admin/orders -> /admin
- /ticket -> /query
- /tickets -> /query
- /new -> /submit
- /create -> /submit
- /form -> /submit
