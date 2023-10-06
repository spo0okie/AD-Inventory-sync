# AD-Inventory-sync
Скрипт односторонней синхронизации БД инвентаризации с Active Directory

При синхронизации пользователя ищет соответствие в БД Инвентори сначала по GUID (записывается после первой синхронизации) логину, затем по ФИО  

## Синхронизирует поля AD -> БД:
* ФИО
* Login
* E-Mail
* GUID (для дальнейшей синхронизации если пользователь будет переименован)
* Должность
* Организация (должна существовать в АД)
* Подразделение (если организация существует в АД)
* Мобильный
* Внутренний номер телефона
* Признак увольнения


Инклудит файл ..\config.priv.ps1 следующего содержания
```powershell
#подразделение где хранятся пользователи
$u_OUDN="OU=Пользователи,DC=domain,DC=local"
#адрес REST API инвентаризации
$inventory_RESTapi_URL="http://inventory.domain.local/web/api"

#писать ли изменения в БД инвентаризации
$write_inventory=$false

#логфайл
$logfile="C:\Joker\Works\PS\ad-inventory-sync\user.log"
```

## Установка
Предположим что мы находимся в папке где уже есть файл config.priv.ps1

```cmd
mkdir libs.ps1
cd libs.ps1
git clone https://github.com/spo0okie/ps1.libs.git .
cd ..
mkdir ad-inventory-sync\user
cd ad-inventory-sync\user
git clone https://github.com/spo0okie/AD-Inventory-sync.git .
ad-to-inventory.cmd
```

## История изменений
v1.0 Initial commit 
     На базе синхронизации Inventory->AD

