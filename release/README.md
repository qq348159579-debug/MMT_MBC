# MBC 工时与物料管理工具发布文件

## 文件

- `MBC工时与物料管理工具.html`：由当前源码构建出的单文件前端，可直接双击打开，也可放到静态服务器。

## 直接使用 HTML

适合只用本地文件/浏览器模式：

1. 打开 `MBC工时与物料管理工具.html`
2. 默认用户密码仍按系统内置数据使用
3. 如需要多人共享 SQL Server 数据，请使用下面的服务器部署方式

## 服务器部署

项目根目录已经提供：

- `Dockerfile`
- `docker-compose.yml`
- `.env.example`
- `index.js`

部署步骤：

```bash
cp .env.example .env
# 修改 .env 中 SQL_SERVER / SQL_DATABASE / SQL_USER / SQL_PASSWORD
docker compose up -d --build
```

访问：

```text
http://服务器IP:3000
```

如果不用 Docker，也可以：

```bash
npm ci
npm run build
npm start
```
