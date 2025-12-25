FROM python:3.9-slim

# 设置中文环境支持（可选，主要针对文件系统查看）
ENV LANG C.UTF-8
ENV LC_ALL C.UTF-8

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple

COPY app.py .

RUN mkdir uploads results

EXPOSE 5000

CMD ["python", "app.py"]
