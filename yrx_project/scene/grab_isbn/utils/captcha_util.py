import cv2
import ddddocr
import numpy as np


def slider_captcha(captcha_bg, captcha_target):
    # 将二进制的图片数据转换为numpy数组
    nparr = np.fromstring(captcha_bg, np.uint8)
    # 将numpy数组转换为OpenCV的图像格式
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)

    # 同样的方式转换模板图像
    template_nparr = np.fromstring(captcha_target, np.uint8)
    template = cv2.imdecode(template_nparr, cv2.IMREAD_COLOR)

    # 使用模板匹配找到暗区域
    result = cv2.matchTemplate(img, template, cv2.TM_CCOEFF_NORMED)

    # 找到匹配度最高的位置
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

    # 计算暗区域的中心坐标
    h, w = template.shape[:2]
    cX = max_loc[0] + w // 2
    cY = max_loc[1] + h // 2


    slide = ddddocr.DdddOcr(det=False, ocr=False)
    res = slide.slide_match(target, bg, simple_target=True)
    lt_x, lt_y, rb_x, rb_y = res.get("target")
    center = (lt_x + rb_x) // 2, (lt_y + rb_y) // 2
    print(res)
    print(center)
    # 在图像上画一个圆来标记中心点
    cv2.circle(img, center, 10, (0, 255, 0), -1)

    top_left = max_loc
    bottom_right = (top_left[0] + w, top_left[1] + h)
    cv2.rectangle(img, top_left, bottom_right, (0, 0, 255), 2)

    # 展示图像
    cv2.imshow('Image', img)
    cv2.waitKey(0)
    cv2.destroyAllWindows()

    return cX, cY


def find_dark_area(image_binary):
    # 将二进制的图片数据转换为numpy数组
    nparr = np.fromstring(image_binary, np.uint8)
    # 将numpy数组转换为OpenCV的图像格式
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)

    # 设置RGB的阈值，你可以根据实际情况调整这个值
    threshold = [50, 50, 50]

    # 找到RGB值都大于阈值的像素
    bright_pixels = cv2.inRange(img, np.array([0, 0, 0]), np.array(threshold))

    # 将这些像素标记为红色
    img[bright_pixels > 0] = [0, 0, 255]

    # 展示图像
    cv2.imshow('Image', img)
    cv2.waitKey(0)
    cv2.destroyAllWindows()


if __name__ == '__main__':
    with open('/bg3.jpg', 'rb') as f:
        bg = f.read()
    with open('/target3.jpeg', 'rb') as f:
        target = f.read()
    slider_captcha(bg, target)
    # find_dark_area(bg)


    print()