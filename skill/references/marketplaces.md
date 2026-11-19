# 站点与邮编字典（15 站）

域名与配送邮编由**站点键**决定。表名/文件名里的前缀决定站点键（如 `US-B09WXSPDQ9(20)__` → `US`，`DE-B0...` → `DE`）。

| 键 | 国家 | 域名 | 配送邮编 |
|---|---|---|---|
| US | 美国 | amazon.com | 90001 |
| UK | 英国 | amazon.co.uk | DN11 0GU |
| DE | 德国 | amazon.de | 10115 |
| JP | 日本 | amazon.co.jp | 001-0024 |
| CA | 加拿大 | amazon.ca | M4Y 1M7 |
| AU | 澳大利亚 | amazon.com.au | 3175 |
| FR | 法国 | amazon.fr | 75020 |
| ES | 西班牙 | amazon.es | 35600 |
| IT | 意大利 | amazon.it | 29015 |
| SG | 新加坡 | amazon.sg | 608829 |
| SE | 瑞典 | amazon.se | 10272 |
| PL | 波兰 | amazon.pl | 01-541 |
| TR | 土耳其 | amazon.com.tr | 06000 |
| IE | 爱尔兰 | amazon.co.uk（同英国域名） | D24 E0CX |
| MX | 墨西哥 | amazon.com.mx | 11000 |

## 别名

`USA/AMZUS` → US，`GB/GBR` → UK，`GER` → DE，`JPN` → JP，`CAN` → CA，`AUS` → AU，`FRA` → FR，`ESP` → ES，`ITA` → IT，`SGP` → SG，`SWE` → SE，`POL` → PL，`TUR` → TR，`IRL` → IE，`MEX` → MX。域名里的 `com/co.uk/de/...` 也参与识别。

## 硬规则

1. **识别不出站点就停**，禁止默认美国。站点错 = 域名错 + 邮编错 + 整份数据作废。
2. 站点确定后，第一个关键词之前必须把配送地址改成该站标准邮编，并读取页面顶部确认文案（如 `Los Angeles 90001`）。确认不到就不开跑。
3. 爱尔兰用 `amazon.co.uk` 域名，但邮编必须是爱尔兰的 `D24 E0CX`。
4. 改了邮编再搜，排名/广告/展示口径才一致；同一批里不要中途反复切换站点或邮编。
