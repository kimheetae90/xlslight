# xlslight

엑셀에서 value 혹은 formula등 필요한 부분만 간소화하여 텍스트로 저장하는 포멧 개발

## NPOI를 사용
NPOI : xlsx를 파싱해주는 라이브러리 중 강력하면서도 무료인 것(Apache-2.0 license)

https://github.com/nissl-lab/npoi/


## YamlDotNet을 사용
xlslight는 yaml 형식을 사용합니다. yaml을 파싱하기 위해 YamlDotNet를 사용

https://github.com/aaubry/YamlDotNet


## xlslight 포멧

* yaml 형식을 사용

* 아래와 같은 구조를 갖음
_추가될 때 마다 문서 최신화 필요_

```
sheets : 각 시트 표기
  name : 시트 이름
  cells : 각 셀 정보
    Offset : 이 전 셀로부터 얼만큼 떨어졌는가, 앞자리 column, 뒷자리 row로 표기하고 앞자리만 있다면 column만 띄워짐
    Type : 셀의 타입으로 NPOI에 선언된 CellType을 표기
    Value : 셀에 기입된 값
```

```
Ex)

sheets:
    - name : sheet1
      cells :
        - Offset : 1,1
          Type : 2
          Value : A3*B3
        - Type : 1
          Value : Hello
    .
    .
    .

    - name : sheet2
      cells :
        - Offset : 0
          Type : 0
          Value : 2
.
.
.
```


