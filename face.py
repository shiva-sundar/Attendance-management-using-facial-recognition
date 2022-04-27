import face_recognition
import cv2
from PIL import Image


def capture_image():
    cam = cv2.VideoCapture(0)
    attendance_image=""

    cv2.namedWindow("Take Image")

    while True:
        ret, frame = cam.read()
        if not ret:
            print("failed to grab frame") 
            break
        cv2.imshow("Take Image", frame)

        key = cv2.waitKey(1)
        #print(key)
        if key== 27 or key==113:
            # ESC pressed
            print("Escape or Quit hit, closing...")
            break
        elif key == 32:
            # SPACE pressed
        
        
            attendance_image = r"C:\Users\bsskt\OneDrive\Desktop\projects\Attendance\temp\temp.jpg"
            cv2.imwrite(attendance_image, frame)
            print("{} written!".format(attendance_image))

    cam.release()

    cv2.destroyAllWindows()
    
    return attendance_image
    
    
    '''
    
def upload_image(path):
    global attendance_image = cv2.imread(path) '''
 




def extract_faces():
    image = face_recognition.load_image_file(r"{}".format(capture_image()))                                                 #r"C:\Users\bsskt\OneDrive\Desktop\projects\Attendance\test3.jpg"
    face_locations = face_recognition.face_locations(image)
    print(face_locations)
    
    extracted_faces_image_names=set()
    
    
    for i, face_location in enumerate(face_locations):
        # Print the location of each face in this image
        top, right, bottom, left = face_location
        
        # We can extract face using these 4 points
        face_image = image[top:bottom, left:right]
        pil_image = Image.fromarray(face_image)
        
        # Show extracted face
        #pil_image.show()
        
        image_name="face-"+str(i)+".jpg"
        extracted_faces_image_names.add(image_name)
        saved_images="C:\\Users\\bsskt\\OneDrive\\Desktop\\projects\\Attendance\\extracted_faces\\" + image_name
        pil_image.save(saved_images)
    
    
    total_extracted_images = len(extracted_faces_image_names)
    print(total_extracted_images)
    return total_extracted_images
    
    
