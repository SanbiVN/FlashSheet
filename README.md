# FlashSheet
Di chuyển tức thời qua lại các trang tính Excel nhanh chóng

[Nhấn tải FlashSheet](https://github.com/SanbiVN/FlashSheet/releases/download/FlashSheet/FlashSheet_v1.3.zip)
[![Lượt tải](https://img.shields.io/github/downloads/SanbiVN/FlashSheet/total.svg)](https://github.com/SanbiVN/FlashSheet/releases/download/FlashSheet/FlashSheet_v1.3.zip) 
Mật khẩu VBA là 1

Add-in với chức năng di chuyển tức thời giữa các trang tính bằng cách rê chuột trên Thanh hiển thị tên trang tính, mà không cần phải thực hiện nhấn vào tên trang tính, giúp chuyển trang tính nhanh hơn, nếu như tệp Excel của bạn có rất nhiều trang tính, thì add-in sẽ thực hiện tốt chức năng của nó.

### Các chức năng của Add-in:

- Tự động chuyển trang tính khi rê chuột trên thanh hiển thị tên.
- Tự động trở lại trang tính hiện hành trước đó.
- Tự động kéo đến các trang sau hoặc trước khi có rất nhiều trang tính.


(Hình ảnh mình họa thanh hiển thị tên trang tính)

![image](https://github.com/SanbiVN/FlashSheet/assets/58664571/68089cd6-9c1a-4353-9cae-df4e64dbdbd2)

​Để tự động kéo đến các trang sau hoặc trước khi có rất nhiều trang tính, rê chuột vào vùng tô màu cam.

### CÁC HÀM THỰC THI

Thực thi	| Hàm	 | Giá trị của tham số	
-------| --------- | -----------
Mở |	=FlashSheet_On()	 |
Tắt	| =FlashSheet_Off()	 |
Độ trễ chuyển trang tính mặc định 30		| =FlashSheet_Delay(30)	 | mili giây từ 20 đến 120	
Đặt lại mặc định	|=FlashSheet_Reset()	 |
Đóng Add-in	| =FlashSheet_Quit()	|
