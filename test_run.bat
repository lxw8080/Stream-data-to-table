@echo off
chcp 65001 >nul
echo ========================================
echo 客户类型标记工具 - 测试流程
echo ========================================
echo.

echo 步骤1: 生成测试数据
python test_marker.py
echo.

echo 步骤2: 运行标记工具
python mark_customer_type.py "测试_租机登记表.xlsx" -o "测试结果.xlsx"
echo.

echo ========================================
echo 测试完成！
echo ========================================
echo.
echo 请打开 "测试结果.xlsx" 查看标记结果
echo.
pause

