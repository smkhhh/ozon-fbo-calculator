#!/bin/bash

cd /home/smkh/ozon-fbo-calculator
git pull origin main

# Если у вас другая ветка, замените main на имя вашей ветки

sudo systemctl restart calculator

