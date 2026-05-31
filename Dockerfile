# ── E-자산통합관리 - Fly.io 배포용 Dockerfile ──
FROM node:20-slim

# better-sqlite3 네이티브 모듈 컴파일에 필요한 빌드 도구
RUN apt-get update && \
    apt-get install -y python3 make g++ && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

# 패키지 먼저 복사 후 설치 (캐시 활용)
COPY package*.json ./
RUN npm install --production

# 소스 코드 복사
COPY . .

# SQLite DB 저장용 데이터 디렉토리 생성
RUN mkdir -p /app/data

EXPOSE 3000

CMD ["node", "app.js"]
