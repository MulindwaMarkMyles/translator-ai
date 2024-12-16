import tarfile
with tarfile.open('boto3-1.35.81-py3-none-any.whl', 'r:gz') as tar:
    tar.extractall('packages/')
