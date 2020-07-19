# Portfolio_Fill-factor-calculation-program

# UPTO5 프로그램 소개
upto5 프로그램은 plastic 실험기기 제조 업체에서 생성되는 데이터를 가공해 fill factor(압력구간)을 찾는 프로그램입니다.

python기반으로 실험데이터를 가공해 excel file로 저장이 가능합니다.

프로그램은 .exe형식으로 구현되었고 python을 전혀 모르더라도 쉽게 데이터를 한눈에 확인하고 가공할 수 있습니다.

실험 데이터는 최대 5개의 file까지 선택이 가능합니다.

# 프로그램 목적
실험의 목적 중 하나인 fill factor(압력구간)를 찾는 일은 시간이 많이 소요됩니다.

실험 데이터는 적게는 수백개의 데이터에서 많게는 수천개의 데이터의 형태를 확인해야하기 때문입니다.

![image](https://user-images.githubusercontent.com/59601177/71885221-3fd7c600-317d-11ea-88a7-ac198d373bf1.png)
##### 가공되지 않은 데이터의 모습

upto5는 실험 데이터를 한 눈에 확인하게 도와줄 뿐만 아니라 fill factor계산까지 자동화한다는 점에서 의의가 있습니다.

# UPTO5 실행과정
1. 프로그램 실행 후 안내문 

![image](https://user-images.githubusercontent.com/59601177/71886686-fccb2200-317f-11ea-9a37-8c8e707c65f1.png)

2. file 선택창을 불러옵니다. (예시 5개의 file을 지정)

![image](https://user-images.githubusercontent.com/59601177/71888680-fe96e480-3183-11ea-9fd1-61b76fed0933.png)

3. 데이터 시각화 및 저장될 excel file명 지정 (예시 test.xlsx)

![image](https://user-images.githubusercontent.com/59601177/71888739-1c644980-3184-11ea-90b4-c45fc84361bc.png)

file명을 입력하면 reference point를 찾으라는 안내가 나타납니다.

![image](https://user-images.githubusercontent.com/59601177/71888886-546b8c80-3184-11ea-9bc0-22cc207d35f9.png)

이때 아래의 돋보기 버튼을 이용해 그래프를 확대할 수 있습니다.

![image](https://user-images.githubusercontent.com/59601177/71888968-87ae1b80-3184-11ea-9fbb-bf0379d9ce43.png)

이후 그래프의 변곡점 지점을 마우스로 가르키면 우측 하단에 x값을 얻을 수 있습니다. 

이때 최초의 변곡점과 이후 4개의 굴곡의 x값을 기록해둬야합니다.

4. 이후 나타나는 창에 확인한 ff-reference값을 입력합니다. (예시 215)

![image](https://user-images.githubusercontent.com/59601177/71887977-8bd93980-3182-11ea-875c-a7dd121a0879.png)

추가로 4개의 reference point를 입력합니다. (예시 220 227 233 240)

![image](https://user-images.githubusercontent.com/59601177/71889072-b88e5080-3184-11ea-8dad-fe4751917308.png)

5. 프로그램은 fill factor를 계산한 뒤 그 결과를 두 종류의 그래프를 통해 보여줍니다. 

![image](https://user-images.githubusercontent.com/59601177/71889108-c9d75d00-3184-11ea-8a29-4433f8812cf5.png)
![image](https://user-images.githubusercontent.com/59601177/71889123-d2c82e80-3184-11ea-9160-577230cb24cf.png)

6. 최종적으로 요약된 결과값을 출력하며 그 결과를 excel file로 저장할 수 있습니다. (예시 test_final.xlsx)

![image](https://user-images.githubusercontent.com/59601177/71889302-32263e80-3185-11ea-9970-5b843a40204a.png)

**excel 형태로 기록된 모습**

![image](https://user-images.githubusercontent.com/59601177/71889379-5aae3880-3185-11ea-862f-c8579362af74.png)

### 사용된 Library
**openpyxl**,
**tkinter**,
**PySimpleGUI**,
**pandas**,
**numpy**,
**matplotlib**
