FROM gitpod/workspace-full-vnc
USER root
RUN sudo apt update && sudo apt install -y python-tk python3-tk tk-dev tk
