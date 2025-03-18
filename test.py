import requests
import threading

# 定义请求的URL和参数
url = "http://localhost:8255/networkForward/pad/saveBloodSampleApi"
headers = {
    "Authorization": "Basic c3dvcmQ6c3dvcmRfc2VjcmV0",
    "Blade-Auth": "bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ0ZW5hbnRfaWQiOiI2MzE0NDQiLCJhY2NvdW50X3R5cGUiOiIxIiwiY29tcGFueV9pZCI6MTY0NTU5Nzc3Mjc2NjgzNDY4OSwidXNlcl9uYW1lIjoiY3IiLCJyZWFsX25hbWUiOiJhZG1pbiIsImF2YXRhciI6Imh0dHA6Ly8xOTIuMTY4LjE2OC4yMDU6OTgwMC9waW1zL3VwbG9hZC8yMDIzMDkxMi81OGRlNDcwN2I1ODk4ZTJmNzI1YjA4OWFjMzRmOWYxZi5qcGciLCJhdXRob3JpdGllcyI6WyJhZG1pbiIsIuermemVvyJdLCJjbGllbnRfaWQiOiJzYWJlciIsInJvbGVfbmFtZSI6IuermemVvyxhZG1pbiIsImxpY2Vuc2UiOiJwb3dlcmVkIGJ5IGJsYWRleCIsInBvc3RfaWQiOiIxNjQ1NTk5MDgwOTc0NDQ2NTk0IiwiY29tcGFueV9zZXJpYWwiOiIxMDMiLCJ1c2VyX2lkIjoiMTY0NTU5Njk1ODAyNzQ3Njk5MyIsInJvbGVfaWQiOiIxMDMyMDIzMDA5MDU0NDQyLDE2NDU1OTY5MjY5MTQxMjk5MjEiLCJzY29wZSI6WyJhbGwiXSwibmlja19uYW1lIjoiYWRtaW4iLCJjb21wYW55X25hbWUiOiLltIfku4HmtYbnq5kiLCJvYXV0aF9pZCI6IiIsImRldGFpbCI6eyJ0eXBlIjoid2ViIn0sImV4cCI6MTczOTE5NjMxOSwiZGVwdF9pZCI6IjE2NDU1OTY5NTcxMTczMTMwMjYiLCJqdGkiOiI0YjUyZmVkMC0zZmNmLTRjMDgtOWE5ZS04MDEwYjU0OTY3MjkiLCJhY2NvdW50IjoiY3IifQ.Yme3zKYiM4skXZ5PCLe4-iVxLkHjxUIOJq0utYCVnUk",
}
data = {
    "donorId": "103L0002675",
    "qualityStatus": "01",
    "status": "",
    "materialSchemeId": "1723260908633911297",
    "sampleType": "2",
    "realName": "cr",
    "realUser": "1645596958027476993",
    "collectPlaceName": "巴山镇六点半小学",
}

data1 = {
    "donorId": "103L0002637",
    "qualityStatus": "01",
    "status": "",
    "materialSchemeId": "1723260908633911297",
    "sampleType": "2",
    "realName": "cr",
    "realUser": "1645596958027476993",
    "collectPlaceName": "巴山镇六点半小学",
}


def send_post_request():
    response = requests.post(url, headers=headers, json=data)
    print(f"Response Status Code: {response.status_code}")
    print(f"Response Body: {response.json()}")


def send_post_request1():
    response = requests.post(url, headers=headers, json=data1)
    print(f"Response Status Code: {response.status_code}")
    print(f"Response Body: {response.json()}")


# 创建两个线程来同时发送请求
thread1 = threading.Thread(target=send_post_request)
thread2 = threading.Thread(target=send_post_request1)

# 启动线程
thread1.start()
thread2.start()

# 等待线程完成
thread1.join()
thread2.join()
