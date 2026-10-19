# LibreOffice Impress and poppler for the shape check of slidev-pptx.
FROM ubuntu:22.04
ENV DEBIAN_FRONTEND=noninteractive
RUN apt-get update && apt-get install -y --no-install-recommends \
      libreoffice-impress poppler-utils fontconfig \
      fonts-liberation fonts-liberation2 fonts-dejavu-core fonts-crosextra-carlito fonts-crosextra-caladea fonts-nanum \
    && rm -rf /var/lib/apt/lists/*
