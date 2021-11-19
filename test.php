<?php
// 一个php调用输出的例子
$ret = shell_exec("C:/Windows/System32/excel2txt.exe ./ex.xlsx");

print_r($ret);