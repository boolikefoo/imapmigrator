my_list = [b'"&BCIENQRBBEIEPgQyBDAETw-1"', b'Drafts', b'"Drafts|template"', b'INBOX', b'Outbox', b'Sent', b'Spam', b'Trash']

new_list = {}
for i, j in enumerate(my_list, start=1):
    print(f'{i}. папка {j}')
    new_list[i] = j

delete_list = list(input("Укажите номера папок для исключения из копирования: ").split())

if not delete_list:
    print('Будут скопированы все папки!')
else:
    print('Данные папки не будут скопированы')
    for i in delete_list:
        print(new_list.get(int(i)))
        new_list.pop(int(i))

print(new_list.values())