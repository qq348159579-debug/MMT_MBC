# MBC 工时与物料管理工具部署包

本目录是可直接交付给内网服务器的静态部署包。

## 文件说明

- `MBC工时与物料管理工具.html`：完整单文件应用，可直接双击打开，也可放到任意静态 Web 服务器。
- `nginx.conf`：Nginx 静态服务配置示例。
- `Dockerfile`：基于 Nginx 的容器部署文件。
- `docker-compose.yml`：一条命令启动容器服务。
- `serve.sh`：已有 Node.js 环境时的本机/服务器简易启动脚本。

## 最简单用法

直接把 `MBC工时与物料管理工具.html` 发给部门用户，双击打开即可使用。

> 当前版本数据保存在浏览器 localStorage 中。更换电脑或浏览器前，请在工具内使用“导入/导出”页面导出 JSON 备份。

## Nginx 部署

1. 将 `MBC工时与物料管理工具.html` 上传到服务器目录，例如：

   `/var/www/mbc/MBC工时与物料管理工具.html`

2. 将 `nginx.conf` 中的 `root /usr/share/nginx/html;` 改为实际目录，例如：

   `root /var/www/mbc;`

3. 将配置放入 Nginx 站点配置目录并 reload：

   ```bash
   sudo nginx -t
   sudo systemctl reload nginx
   ```

4. 内网访问：

   `http://服务器IP/`

## Docker 部署

在本目录执行：

```bash
docker compose up -d --build
```

默认端口：

`http://服务器IP:8080/`

如需改端口，编辑 `docker-compose.yml` 的 `ports`。

## Node.js 简易静态服务

已有 Node.js 时，在本目录执行：

```bash
chmod +x serve.sh
./serve.sh
```

默认端口：

`http://服务器IP:8080/`
