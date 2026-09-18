# K.S. Distillery Report Transformer
#
# Python 3.11 to match the environment the build scripts have been running on.
FROM python:3.11-slim

# LibreOffice is required by the PDF-producing scripts (bond drill-downs, pace
# PDFs). Omit it and those builds fail at the point of export rather than at
# startup, so it is installed even though the v1 streams do not use it.
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-calc \
        fonts-dejavu \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /srv

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app ./app
COPY scripts ./scripts
# The PDF builders this service owns (bond cumulative books, the comparative
# page, the liquidation scorecard). They are called by path, so they have to be
# in the image, not only in the repo.
COPY reports ./reports
# Reference material for a disk that comes up bare: the 81 build scripts and
# the master data. Copied into the workspace on first boot, never overwritten.
COPY seed ./seed
COPY verify_pipeline.py .

# The workspace is a MOUNTED VOLUME, never baked into the image. The workbooks
# are accumulated state: a month's earlier days exist only inside them.
ENV KSD_WORKSPACE_ROOT=/data/workspace
VOLUME ["/data"]

EXPOSE 8000

# One worker. The pipelines mutate shared workbook state in fixed folders, so
# a second worker process would interleave writes to the same files.
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "1"]
