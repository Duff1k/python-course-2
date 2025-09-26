# file = open("demo_read.txt", encoding="utf-8")
# contents = file.read()
# print(contents)
# file.close()


# with open("demo_read.txt", "r" ,encoding="utf-8") as file:
#     lines = file.readlines()
#     words = lines[1].split()
#     print(type(words))
#     print(".".join(words[1:]))

#
# with open("demo_read.txt", "r" ,encoding="utf-8") as file:
#     for i, line in enumerate(file, start=1):
#         if i > 5:
#             break
#         print(line.strip())

# Вывод со второй строки по 4 строку и в каждой строке выводим в обратном направлении с заменой пробела на символ /
# with open("demo_read.txt", "r", encoding="utf-8") as file:
#     for i, line in enumerate(file, start=1):
#         if 2<= i <=4:
#             line_without_n = line.strip()
#             joined = "/".join(line_without_n.split())
#             reversed_line = joined[::-1]
#             print(reversed_line)
#         if i > 4:
#             break


# with open("demo_read.txt", "r", encoding="utf-8") as file:
#     for line in file:
#         clean_line = line.strip()
#         if clean_line.startswith("П"):
#             print(clean_line)

# with open("demo_read.txt", "r", encoding="utf-8") as file:
#     for line in file:
#         if line.strip():
#             print(line.strip())

# MIN_LEN = 5
#
# with open("demo_read.txt", "r", encoding="utf-8") as file:
#     for line in file:
#         line = line.strip()
#         if len(line) > MIN_LEN:
#             print(line)
#
#
#
# with open("demo_read.txt", "r", encoding="utf-8") as file:
#     for line in file:
#         line = line.strip()
#         if int(line)%2==0:
#             print(line)

# with open("demo_write.txt", "w", encoding="utf-8") as file:
#     file.write("Привет мир 1!\n")
#     file.write("Привет мир 2!\n")
#     file.write("Привет мир 3!\n")


# lines = ["Первая строка\n", "Вторая строка\n", "Третья строка\n"]
#
# with open("demo_write.txt", "w", encoding="utf-8") as file:
#     file.writelines(lines)
#
# with open("demo_student_numb.txt", "w", encoding="utf-8") as file:
#     for i in range(1, 101):
#         if i % 2 == 0:
#             continue
#         file.write(str(i) + "." + "\n")

# with open("demo_read.txt", "r", encoding="utf-8") as file:
#     lines = file.readlines()
#
#
# target_index = 2
# if 1<=target_index<=len(lines):
#     lines[target_index - 1] = "Как7!!!\n"
#
# with open("demo_read.txt", "w", encoding="utf-8") as file:
#     file.writelines(lines)

# with open("demo_write.txt", "r", encoding="utf-8") as file:
#     text = file.read()
#
# text = text.replace("строка", "строчка")
#
# with open("demo_write.txt", "w", encoding="utf-8") as file:
#     file.write(text)

# start_line, end_line = 6, 9
#
# with open("demo_write.txt", "r", encoding="utf-8") as file:
#     lines = file.readlines()
#
# lines = lines[start_line-1:end_line]
#
# with open("demo_write.txt", "w", encoding="utf-8") as file:
#     file.writelines(lines)


# student = ["Аня\n", "Катя\n", "Ваня\n", "Коля\n"]
#
# with open("demo_append.txt", "a", encoding="utf-8") as file:
#     file.writelines(student)

with open("demo_write.txt", "r", encoding="utf-8") as file:
    lines = file.readlines()

insert_pos = 2

lines.insert(insert_pos-1, "Вставленная строка\n")

with open("demo_write.txt", "w", encoding="utf-8") as file:
    file.writelines(lines)





