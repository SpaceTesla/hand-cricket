FROM gitpod/workspace-full-vnc
USER root
RUN apt update && apt install -y python-tk python3-tk tk-dev
