@echo off
@chcp 65001>nul
title msoxfix script by Jose Luis Rincon Ciriano - ©2021

color 3F
@echo:
for %%a in (*.msox) do (
    @echo Fixing %%a, please wait...
    msoxfix %%a
    @echo:
)
@echo:
pause