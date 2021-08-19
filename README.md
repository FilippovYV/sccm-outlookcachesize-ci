# sccm-outlookcachesize-ci

Скрипты для управления размером локального кэша для Outlook 2016/2019/365.

По умолчанию Outlook выбирает размер кэша (глубину хранения) в зависимости от размера диска.

Hard drives up to 32Gb in size, default to 1 month

Hard drives bigger than 32GB, but less than 64 GB, default to 3 months

Hard drives 64GB and larger, default to 12 months

Это хорошо работает, если на компьютере только один пользователь, но если пользователей 5-10, то на 256GB SSD диске всё место уходит под кэш.
Можно отключить Exchange Cached Mode, но тогда пользовательский опыт изрядно ухудшается. Мне нужен был способ управления размером кэша.

Для этого есть административные шаблоны групповых политик, также можно напрямую править реестр в [Software\Policies], но этот метод не очень хороший.
Обычно, если у нас не закачан в кэш весь почтовый ящик целиком, то можно получить дополнительные данные нажав на ссылку возле последнего сообщения - 
Click here to view more on Microsoft Exchange. Но если глубина хранения ограничена политиками, то этой ссылки нет, дополнительные данные получить невозможно.
От этого получается разрыв - например, политика архивации настроена на 1 год, т.е. все письма старше года уходят в архив. Если при этом ограничить глубину хранения, например, в 3 месяца,
то мы не можем в Outlook увидеть письма, в периоде от 3-х месяцев до года. Эта проблема давняя и она до сих пор не решена.

При этом, если регулировать глубину хранения непосредственно через интерфейс клиента, такой проблемы не возникает. Эксперименты и поиски места хранения этой настройки показали, 
что глубина хранения регулируется двоичным ключом в реестре, в настройках профиля.

Это ключ с именем 00036649, находится в настройках профиля в реестре и хранит размер кэша в месяцах. Например, три месяца будут храниться как 03 00 00 00. Шесть месяцев как 06 00 00 00.
Особое значение 00 0 00 00 означает "весь яшик". Сами профили находятся в разделе реестра HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\.

Посмотреть настройки кэша для текущего пользователя можно так:
````powershell
$EmailToKeepRegistry = '00036649'
$Outlook16Profiles = 'HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Profiles\'
Get-ChildItem -Path Registry::$Outlook16Profiles -Recurse | Where-Object { $_.Property -eq $EmailToKeepRegistry } | ForEach-Object { (Get-ItemProperty -Path Registry::$_ -Name $EmailToKeepRegistry) }
````
Значение вида 00036649 : {XX, 0, 0, 0} покажет глубину хранения в XX месяцев.

На основе этого сделаны Configuration Item/Baseline для ECM(ConfigMgr).

Скрипт OutlookCacheSizeDetection.ps1 проверяет настройки пользователя и возвращает $True если настройки правильные и $False если глубина хранения отличается от заданной.
Скрипт OutlookCacheSizeRemediate.ps1 прописывает нужные настройки в реестр для текущего пользователя.

В обоих скриптах переменная $EmailToKeepMonths задаёт глубину хранения в месяцах.

Как это работает - Outlook перечитывает это значение при запуске, потом если глубина хранения выросла, то Outlook начинает закачивать в кэш дополнительные письма.
Если же глубина уменьшилась, то Outlook начинает в фоновом режиме перемещать (дефрагментировать) кэш и урезать его. Процесс приостанавливается при любой другой активности в кэшем, например,
отправка/прием писем.

Baseline можно назначить как на пользователей, так и на компьютеры. Во втором случае настройка будет применена ко всем пользователя компьютера, по мере их входа.




