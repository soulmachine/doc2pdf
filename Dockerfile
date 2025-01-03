# See https://playwright.dev/docs/docker
FROM node:20-bookworm

# Install Playwright with dependencies
RUN npx -y playwright@1.49.1 install --with-deps

# Copy application files
COPY requirements.txt doc2pdf.py /app/

# Set working directory
WORKDIR /app

# Install Python environment and dependencies
RUN apt -y update && \
    apt install -y \
        python3-pip && \
    pip3 install --break-system-packages -r requirements.txt

# Set default command
ENTRYPOINT ["python3", "/app/doc2pdf.py"]
