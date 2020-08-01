import cv2
import os

cap = cv2.VideoCapture(0)
cnt = 0
print("Enter your name : ")
newFolder = str(input())
directory = "../main/images/"
path = os.path.join(directory, newFolder)
os.mkdir(path)

while True:
    ret, image = cap.read(0)

    cv2.imwrite(path + "/img_" + str(cnt) + ".jpg", image)
    cnt += 1
    print('Capturing image...')
    cv2.imshow("Capturing image", image)
    cv2.waitKey(1)

    if cnt == 50:
        break


cap.release()
cv2.destroyAllWindows()


