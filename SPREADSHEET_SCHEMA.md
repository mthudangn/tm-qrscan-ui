# Spreadsheet Schema

## PXK

Suggested fields:

```text
prodKey
Tên Sp
DH
Customer
Ngay_GTAM
Ngay_CAN
Ngay_IN
Ngay_CHAP
Ngay_DAN
Ngay_QC
Ngay_KHO
_BaseKey
```

## QUEUE

```text
request_id
created_at
sid
pxk
prodKey
step
status
attempt_count
processed_at
error_message
```

## LOG

```text
timestamp
level
event
sid
pxk
prodKey
step
message
```
