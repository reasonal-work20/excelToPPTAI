@echo off
rem ============================================================
rem  Weekly Report Agent - Dashboard launcher
rem  Double-click this file to open the dashboard.
rem ============================================================
start "Weekly Report Dashboard" powershell -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0WeeklyReportDashboard.ps1"
