#!/bin/bash

install_pynmap(){
    pip uninstall -y python-nmap
    cd /tmp
    git clone https://github.com/dogasantos/python-nmap
    cd python-nmap
    python setup.py install
    cd /tmp
    rm -rf python-nmap
}
pip install -r requirements.txt
install_pynmap

