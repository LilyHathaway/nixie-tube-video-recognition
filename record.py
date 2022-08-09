import cv2
import numpy as np
import imutils
from imutils import contours
import os
import xlwt
import sys

model_path = "imgs\\model.jpg"
video_path = "video\\2.mov"
cut_path = 'cut\\'
excel_path = 'excel.xls'
get_fat=20  #把数字轮廓变胖，不然有些太暗的数字线条连不到一块，数值越大越胖，画面越近，数值应调小
height=150  #数字比对阈值，高度比这个值大的判断为数字，单位像素，如果判断的数字数量不对，试着调一下这个


def model(model_path):
    temp = cv2.imread(model_path)
    gray = cv2.cvtColor(temp, cv2.COLOR_BGR2GRAY)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    # edged = cv2.Canny(blurred, 50, 200, 255)
    # cv2.imwrite(dir + 'edge.png', edged)

    # 使用阈值进行二值化
    thresh = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]
    # cv2.imwrite(dir + 'thresh1.png', thresh)
    kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (1, 5))
    # 使用形态学操作进行处理
    thresh = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel)
    # cv2.imwrite(dir + 'thresh2.png', thresh)
    #膨胀
    peng = cv2.getStructuringElement(cv2.MORPH_RECT, (8, 8)) #矩形结构
    dilation = cv2.dilate(thresh, peng)

    refCnts, hierarchy = cv2.findContours(dilation.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)


    # 画出找到的轮廓
    # cv2.drawContours(temp, refCnts, -1, (0, 0, 255), 2)
    # cv2.imwrite('imgs\\temp1.png', temp)
    refCnts = imutils.contours.sort_contours(refCnts, method="left-to-right")[0]    # 排序，从左到右，从上到下
    digits = {}

    # print("模板数量：",len(refCnts))

    # 遍历每一个轮廓
    for (i, c) in enumerate(refCnts):
    # 计算外接矩形并且resize成合适大小
        (x, y, w, h) = cv2.boundingRect(c)
        roi = thresh[y:y + h, x:x + w]
        roi = cv2.resize(roi, (57, 88))

        # 每一个数字对应每一个模板
        digits[i] = roi
    return digits

def scan_img(key,img,digits):
    # img = cv2.imread(img_path)
    # 将输入图片裁剪到固定大小
    img = imutils.resize(img, height=500)
    # 将输入转换为灰度图片
    img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # 进行高斯模糊操作
    img_blurred = cv2.GaussianBlur(img_gray, (5, 5), 0)
    # 执行边缘检测
    # img_edged = cv2.Canny(img_blurred, 50, 200, 255)
    # cv2.imwrite(dir + 'img_edged.png', img_edged)
    # 使用阈值进行二值化
    img_thresh = cv2.threshold(img_blurred, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
    # cv2.imwrite(dir + 'img_thresh'+ str(key) +'.png', img_thresh)
    img_kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (1, 5))
    # 使用形态学操作进行处理
    img_thresh = cv2.morphologyEx(img_thresh, cv2.MORPH_OPEN, img_kernel)
    # cv2.imwrite('cut\\img_thresh'+ str(key) +'.png', img_thresh)
    # 膨胀
    peng = cv2.getStructuringElement(cv2.MORPH_RECT, (get_fat, get_fat)) #矩形结构
    dilation = cv2.dilate(img_thresh, peng)
    
    # 计算轮廓
    threshCnts, hierarchy = cv2.findContours(dilation.copy(), cv2.RETR_EXTERNAL,	cv2.CHAIN_APPROX_SIMPLE)
    # cv2.drawContours(img, threshCnts, -1, (0, 0, 255), 2)
    # cv2.imwrite('cut\\img'+ str(key) +'.png', img)
    
    
    locs = []
    
    # 遍历轮廓
    for (i, c) in enumerate(threshCnts):
        # 计算矩形
        (x, y, w, h) = cv2.boundingRect(c)
    
        if height < h :
            # 符合的留下来
            locs.append((x, y, w, h))
        else:
            cv2.rectangle(img_thresh,(x,y),(x+w,y+h),(0,0,0),-1)
            # cv2.imwrite('cut\\imgsss'+ str(key) +'.png', img_thresh)
    
    # print(locs)
    
    
    # 将符合的轮廓从左到右排序
    locs = sorted(locs, key=lambda o: o[0])
    output = []
    
    
    result = img.copy()
    
    # 遍历每一个轮廓中的数字
    for (i, (gX, gY, gW, gH)) in enumerate(locs):
        # initialize the list of group digits
        groupOutput = []
    
        # 根据坐标提取每一个组
        group = img_thresh[gY :gY + gH + 5, gX :gX + gW + 5]
        # print("x---",gX,"y---",gY,"w---",gW,"h----",gH)
        # cv2.imwrite('cut\\img_each'+ str(key) +'.png', img_thresh)
    
        # 预处理
        group = cv2.threshold(group, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]

        #膨胀
        peng = cv2.getStructuringElement(cv2.MORPH_RECT, (get_fat, get_fat)) #矩形结构
        dilation = cv2.dilate(group, peng)
    
        # 计算每一组的轮廓
        digitCnts, hierarchy = cv2.findContours(dilation.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        digitCnts = contours.sort_contours(digitCnts, method="left-to-right")[0]
    
        # 计算每一组中的每一个数值
        for c in digitCnts:
            # 找到当前数值的轮廓，resize成合适的的大小
            (x, y, w, h) = cv2.boundingRect(c)
            roi = group[y:y + h, x:x + w]
            roi = cv2.resize(roi, (57, 88))
    
            # 计算匹配得分
            scores = []
    
            # 在模板中计算每一个得分
            for (digit, digitROI) in digits.items():
                # 模板匹配
                res = cv2.matchTemplate(roi, digitROI, cv2.TM_CCOEFF)
                (_, score, _, _) = cv2.minMaxLoc(res)
                scores.append(score)
    
            # 得到最合适的数字
            groupOutput.append(str(np.argmax(scores)))
    
        # 画出来
        cv2.rectangle(result, (gX - 5, gY - 5), (gX + gW + 5, gY + gH + 5), (0, 0, 255), 1)
        cv2.putText(result, "".join(groupOutput), (gX, gY - 15), cv2.FONT_HERSHEY_SIMPLEX, 0.65, (0, 0, 255), 2)
    
        # 得到结果
        output.extend(groupOutput)
    
    re = int("{}".format("".join(output)))
    # 打印结果
    # print("number: {}".format("".join(output)))
    # contrast = np.hstack((img, result))
    
    # 保存结果
    cv2.imwrite(cut_path + "result" + str(key+1) + ".jpg", result)
    # cv2.imwrite(dir + "contrast.jpg", contrast)
    return re

def del_files(dir_path):
    if os.path.isfile(dir_path):
        try:
            os.remove(dir_path) # 这个可以删除单个文件，不能删除文件夹
        except BaseException as e:
            print(e)
    elif os.path.isdir(dir_path):
        file_lis = os.listdir(dir_path)
        for file_name in file_lis:
            # if file_name != 'wibot.log':
            tf = os.path.join(dir_path, file_name)
            del_files(tf)

def cut_video(video_path):
    i = 0
    j = 0
    imgs = []
    rval = True
    del_files(cut_path)
    if not os.path.exists(cut_path):
        os.mkdir(cut_path)
    videoCapture = cv2.VideoCapture(video_path)
    # TODO 咋就一直打不开呢？？？
    if videoCapture.isOpened():  # 判断是否正常打开
        # timeF = int(videoCapture.get(cv2.CAP_PROP_FPS)) #视频帧计数间隔频率
        rate = videoCapture.get(5)
        frame_num = videoCapture.get(7)
        duration = int(frame_num / rate)
        # print(duration)
    else:
        return imgs
    while rval:
        rval, frame = videoCapture.read()
        i += 1
        if (i%rate==0):
            imgs.append(frame)
            j += 1
            cur = int(j/duration*100)
            print("\r", end="")
            print("进度: {}%: ".format(cur), "▓" * (cur // 2), end="")
            sys.stdout.flush()
            cv2.waitKey(1) #延时1ms
    videoCapture.release() #释放视频对象
    return imgs

def save_excel(re):
    if os.path.exists(excel_path):
        os.remove(excel_path)
    book = xlwt.Workbook(encoding='utf-8',style_compression=0)
    sheet = book.add_sheet('01',cell_overwrite_ok=True)
    for i in range(len(re)):
        sheet.write(i,0,i+1)
        sheet.write(i,1,re[i])
    book.save(excel_path)

if __name__ == '__main__':
    print("开始视频处理....")
    imgs = cut_video(video_path)
    if len(imgs) == 0:
        print("视频打开失败....")
        sys.exit(0)
    print("\n截图完毕....")
    digits = model(model_path)
    print("模板分析完毕....")
    result = []
    size = len(imgs)
    print("开始图像识别....")
    for key,img in enumerate(imgs):
        result.append(scan_img(key,img,digits))
        cur = int((key+1)/size*100)
        print("\r", end="")
        print("进度: {}%: ".format(cur), "▓" * (cur // 2), end="")
        sys.stdout.flush()
    # print(result)
    save_excel(result)
    print("\n结果保存完毕，路径---{}".format(excel_path))
