# !/bin/bash
pronto_dir=$(dirname "$pwd")
tag="docker_pronto:v1"
docker build -f $pronto_dir/PRONTO_docker/Dockerfile -t $tag .
