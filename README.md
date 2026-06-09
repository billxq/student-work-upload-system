# 学生作品提交系统

这是一个基于 Python 标准库实现的学生作品收集系统，支持：

- 从当前目录下的 `xlsx` 文件读取学生名单
- 使用 `学籍号后 8 位 + 姓名 + 密码` 登录
- 登录后上传作品图片，支持一次选择多张
- 自动按 `班级/姓名/作品类型_序号.jpg` 保存
- 上传完成后展示成功结果和保存文件列表

## 启动方式

在当前目录执行：

```bash
python app.py
```

然后在浏览器打开：

```text
http://127.0.0.1:8000
```

## 默认登录密码

```text
111111
```

## 学生名单 Excel

程序会自动读取当前目录下的 `xlsx` 文件，并优先选择文件名里包含 `学生名单` 的表格。

Excel 表头需要包含以下三列：

- `班级`
- `学生姓名`
- `学生上海市学籍号`

## 上传结果保存位置

图片会保存到：

```text
uploads/班级/姓名/作品类型_1.jpg
uploads/班级/姓名/作品类型_2.jpg
```

上传记录会额外保存到：

```text
data/upload_records.json
```

## 服务器部署

推荐用 `systemd` 管理 Python 服务，用 `Caddy` 对外反代。

### 1. Python 服务

先把项目放到服务器上的固定目录，例如：

```text
/opt/student-work-upload-system
```

然后创建一个 `systemd` 服务文件，例如：

```bash
sudo nano /etc/systemd/system/student-work-upload.service
```

内容如下：

```ini
[Unit]
Description=Student Work Upload System
After=network.target

[Service]
Type=simple
WorkingDirectory=/opt/student-work-upload-system
ExecStart=/usr/bin/python3 /opt/student-work-upload-system/app.py
Restart=always
RestartSec=3
User=www-data
Group=www-data

[Install]
WantedBy=multi-user.target
```

如果你的 `python3` 不在 `/usr/bin/python3`，请改成实际路径。

启用并启动服务：

```bash
sudo systemctl daemon-reload
sudo systemctl enable student-work-upload
sudo systemctl start student-work-upload
sudo systemctl status student-work-upload
```

查看日志：

```bash
journalctl -u student-work-upload -f
```

### 2. Caddy 反代

在 `Caddyfile` 中配置你的域名，例如：

```caddyfile
work.ssfx.org {
	reverse_proxy 127.0.0.1:8000
	encode zstd gzip
}
```

然后重载 Caddy：

```bash
sudo caddy reload
```

这样公网访问就会通过 Caddy 转发到本机的 Python 服务。

## 注意事项

- `xlsx` 文件和 `github_token` 都不会被上传到仓库
- 如果你要在服务器上长期使用，建议把上传目录和日志目录按实际需要做备份
