# Version 1.0
FROM continuumio/miniconda3:23.9.0-0

LABEL maintainer="Xiaoli.Zhang@rr-research.no"

RUN conda update -n base -c defaults conda

RUN conda install xlrd -c conda-forge

RUN conda install xlutils -c conda-forge

RUN conda install python-pptx -c conda-forge

RUN conda install python-docx -c conda-forge

RUN mkdir -p /pronto

CMD ["python", "/pronto/Script/PRONTO.py"]
