FROM python:3.9-slim

WORKDIR /app

COPY requirements.txt requirements.txt
RUN pip install --no-cache-dir -r requirements.txt
RUN pip install --no-cache-dir numpy pandas
RUN pip install xlsxwriter
RUN pip install openai==0.28
RUN pip install selenium

COPY . .

EXPOSE 5000
CMD ["python", "app/main.py"]

