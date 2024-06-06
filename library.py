from pyzbar.pyzbar import decode
import cv2

# Test decoding a sample image
image = cv2.imread('your_image_with_qr_or_barcode.png')  # Replace with your image path
decoded_objects = decode(image)

for obj in decoded_objects:
    print("Type:", obj.type)
    print("Data:", obj.data.decode("utf-8"))
