	Python класс для работы с 1С
	
	Класс позволяющий обращаться к объектами 1С посредством технологии COM

    Пример использования в CreateDoc_ЗаявкаНаВыполнениеРабот()

    Функции работы с объектом Документ
    ----------------------------------
    getDocAtr(doc_name) -- Возвращает список реквизитов документа
    getDocAtr(doc_name) -- Возвращает значение реквизитов документа
    docByNum(doc_name, **kwargs): -- Поиск документа по номеру и дате
    docByAttr(doc_name, **kwargs): -- Поиск документа по реквизиту
    getDocTab(doc_name)  - Возвращает список табличных частей документа с их реквизитами
    пример реализоции в CreateDoc_ЗаявкаНаВыполнениеРабот() -- Создает новый экземпляр документа
    пример реализоции в CreateDoc_ЗаявкаНаВыполнениеРабот() -- Добавление записи в табличную часть документа
    docDelMark(True) --  Пометки документа на удаление. Установить/Снять (True/False)
    docPosted(doc_ref) -- Проведение документа

    Функции работы с объектом Справочник
    ----------------------------------
    getCatalogsAtr(cat_name) -- Возвращает список реквизитов справочника
    getCatalogsVal(cat_ref, atr_name) -- Возвращает значение указанного реквизита записи справочника
    getCatalogsVals(cat_ref) -- Возвращает значения реквизитов записи справочника
    getLinkByAttr(cat_name, **kwargs) -- Возвращает ссылку на запись элемента справочника по его реквизиту
    getLinkByName(st1, st2) -- Возвращает ссылку на запись элемента справочника по его наименованию

    Функции работы с объектом РегистрСведений
    ----------------------------------
    getInformationRegisters(self, reg_name, *args, cn = None) -- Выгрузка значений регистра сведений
    loadFileToReg(self, doc_ref, img_src) -- Загрузка бинарного файла в регистр сведений

    Облок общих функций
    ------------------------
    _currDate() -- Возврашает текущую дату вида '27.05.2020 14:06:39'
    chAccount() -- Проверка наличия лицевого счёта
