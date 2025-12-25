基于python自动excel比对
 使用 Docker Compose 管理（推荐），创建 docker-compose.yml：
md
 services:
  compare-tool:
    image: ghcr.io/你的用户名/你的仓库名:latest
    ports:
      - "80:5000"
    environment:
      - WEB_USER=your_admin_name
      - WEB_PWD=your_secret_password
    restart: always
