"""
для установки pythoncom, нужно выполнить
pip install pywin32
"""
import win32com.client
import time

# coding=cp1251

class Class_1C():
    """Класс позволяющий обращаться к объектами 1С посредством технологии COM

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
    """

    def __init__(self, conn_str, debug=False):
        """Инициализация класса

        :param conn_str: Строка подключения
        :param debug: Режим отладки
        """
        self.V83 = win32com.client.Dispatch("V83.COMConnector").Connect(conn_str)
        if debug: print('1C Connect -- ' + self._currDate())

    def chAccount(self, lsh):
        """Проверка наличия лицевого счёта

        :param lsh: Номер л/сч
        :return: True / False
        """
        cat_ref = self.getLinkByAttr('Абоненты', attr='ЛицевойСчет', value=lsh)
        return bool(cat_ref.ЛицевойСчет)

    def docDelMark(self, doc_ref, mark=True):
        """Пометки документа на удаление. Установить/Снять (True/False)

        :param doc_ref: Ссылка документа
        :param mark: Установить/Снять (True/False)
        :return: True
        """
        doc_obj = doc_ref.GetObject()  # --> DocumentObject
        doc_obj.SetDeletionMark(mark)
        return True

    def docByNum(self, doc_name, **kwargs):
        """Поиск документа по номеру и дате

        :doc_name: Наименование документа
        :param kwargs: number='000000460', data='20.01.2021  8:30:27'
        :return: ДокументСсылка
        """
        doc_ref = getattr(self.V83.Documents, doc_name).FindbyNumber(kwargs['number'], kwargs['data'])
        return doc_ref

    def docByAttr(self, doc_name, **kwargs):
        """Поиск документа по реквизиту

        :name: Наименование документа
        :param kwargs: attr='', value=''
        :return: ДокументСсылка
        """
        doc_ref = getattr(self.V83.Documents, doc_name).FindByAttribute(kwargs['attr'], kwargs['value'])
        return doc_ref

    def getLinkByAttr(self, cat_name, **kwargs):
        """Возвращает ссылку на запись элемента справочника по его реквизиту

        :name: Наименование справочника
        :param kwargs: attr='ЛицевойСчет', value='14799511640'
        :return: СправочникСсылка
        """
        cat_ref = getattr(self.V83.Catalogs, cat_name).FindByAttribute(kwargs['attr'], kwargs['value'])
        return cat_ref

    def _currDate(self):
        """Возврашает текущую дату

        :return: Дата вида '27.05.2020 14:06:39'
        """
        tm_sec = time.time() + 10800
        struct_time = time.gmtime(tm_sec)
        dt = time.strftime('%d.%m.%Y %X', struct_time)
        return dt

    def getLinkByName(self, st1, st2):
        """Возвращает ссылку на запись элемента справочника по его наименованию

        :param st1: Наименование справочника
        :param st2: Значение реквизита "Наименование"
        :return: Ссылка на запись
        """
        q = f'''
        ВЫБРАТЬ 
        Объект.Ссылка КАК Ссылка
        ИЗ {st1} КАК Объект
        ГДЕ Объект.Наименование like "%{st2}%"
        '''
        query = self.V83.NewObject("Query", q)
        sel = query.Execute().Choose()
        ls = []
        while sel.next():
            ls.append(sel)
        return ls[0].Ссылка


    def docPosted(self, doc_ref):
        """Проведение документа

        :param doc_ref: Ссылка документа
        :return: True
        """
        doc_obj = doc_ref.GetObject()
        doc_obj.Posted = True  # Провести документ
        doc_obj.Write()
        return True

    def CreateDoc_ЗаявкаНаВыполнениеРабот(self, **kwargs):
        """Создание документа ЗаявкаНаВыполнениеРабот

        :param kwargs: {lsh: ЛицевойСчет, txt_z: ТекстЗаявки, vid: ВидРабот}
        :return:
        """
        doc_mgr = getattr(self.V83.Documents, "ЗаявкаНаВыполнениеРабот")  # --> DocumentManager
        doc_obj = doc_mgr.CreateDocument()
        cr_dt = self._currDate()

        ab_ref = self.getLinkByAttr('Абоненты', attr='ЛицевойСчет', value=kwargs['lsh'])
        setattr(doc_obj, "Дата", cr_dt)
        setattr(doc_obj, "Организация", self.getLinkByName('Справочник.Организации', 'Газпром межрегионгаз Орёл'))
        setattr(doc_obj, "Подразделение", self.getCatalogsVal(ab_ref, 'Подразделения'))
        setattr(doc_obj, "ТекстЗаявки", kwargs['txt_z'])
        setattr(doc_obj, "Абонент", ab_ref)
        setattr(doc_obj, "ТипЗаявки", self.getLinkByName('Справочник.ТипыЗаявокАбонентов', 'Работы контролеров'))
        setattr(doc_obj, "ПланируемаяДата", cr_dt)

        # Добавление записи в табличную часть документа
        link = self.getLinkByName('Справочник.ВидыДокументовУчетРабот', kwargs['vid'])
        str_tab = doc_obj.СписокРабот.Add()  # Добавляет строку в конец табличной части. Возвращает "Строка табличной части"
        str_tab.ВидРаботы = link
        doc_obj.Posted = True  # Провести документ

        doc_obj.Write()
        # print('Документ создан')
        return doc_obj.Ref

    def getDocAtr(self, doc_name):
        """Возвращает список реквизитов документа

        :param doc_name: Имя документа
        :return: Список реквизитов. Type: <class 'list'>
        """
        doc_ref = getattr(self.V83.Documents, doc_name).Select()  # --> DocumentObject
        doc_ref.next()
        doc_obj = doc_ref.GetObject()  # --> DocumentObject
        doc_md = doc_obj.Metadata()  # --> MetadataObject
        ls_atr = [atr.name for atr in doc_md.Attributes]
        return ls_atr

    def getDocTab(self, doc_name):
        """Возвращает список табличных частей документа с их реквизитами

        :param doc_name: Имя документа
        :return: Словарь реквизитов. Type: dict:[list]
        """
        doc_ref = getattr(self.V83.Documents, doc_name).Select()  # --> DocumentObject
        doc_ref.next()
        doc_obj = doc_ref.GetObject()  # --> DocumentObject
        doc_md = doc_obj.Metadata()  # --> MetadataObject
        # doc_md.TabularSections -- колекция 'Табличных частей'  документа
        dic_atr = {tab.name: [atr.name for atr in tab.Attributes] for tab in doc_md.TabularSections}
        return dic_atr

    def getCatalogsAtr(self, cat_name):
        """Возвращает список реквизитов справочника

        :param cat_name: Имя справочника
        :return: Список реквизитов. Type: <class 'list'>
        """
        cat_sel = getattr(self.V83.Catalogs, cat_name).Select()  # --> CatalogSelection
        cat_sel.next()  # --> bool
        cat_obj = cat_sel.GetObject()  # --> CatalogObject
        cat_md = cat_obj.Metadata()  # --> MetadataObject
        ls_atr = [atr.name for atr in cat_md.Attributes]
        return ls_atr

    def getInformationRegisters(self, reg_name, *args, cn=None):
        """Выгрузка значений регистра сведений

        :param reg_name: Имя регистра сведений
        :param args: Список полей
        :param cn: Кол-во выгружаемых записей
        :return: Результат запроса. Type: <class 'list'>

        !! Добавить обработку типа дата
        """
        if len(args) < 1:
            q = f'ВЫБРАТЬ первые 0 "" КАК col_0 ИЗ РегистрСведений.{reg_name} КАК {reg_name}'
        else:
            cn = 'ПЕРВЫЕ ' + str(cn) if bool(cn) else ''
            str_arg = ''.join([s + ' КАК col_' + str(i) + ', ' for i, s in enumerate(args)])[:-2]
            q = f'''ВЫБРАТЬ {cn} {str_arg} ИЗ РегистрСведений.{reg_name} КАК {reg_name}'''
        ls = []
        query = self.V83.NewObject("Query", q)
        sel = query.Execute().Choose()
        sel = query.Execute().Select()  # -->  ВыборкаИзРезультатаЗапроса
        while sel.next():
            ls_0 = []
            [ls_0.append(sel.Get(i) if type(sel.Get(i)) == str else sel.Get(i).Наименование) for i in range(len(args))]
            ls.append(ls_0)
        return ls


    def getCatalogsVal(self, cat_ref, atr_name):
        """Возвращает значение указанного реквизита записи справочника

        :param cat_ref: Ссылка на запись справочника
        :param atr_name: Наименование реквизита справочника
        :return: Значение реквизитов справочника
        """
        cat_name = cat_ref.Metadata().Name
        ls_attr = self.getCatalogsAtr(cat_name)
        naim = cat_ref.Наименование
        q = f'''ВЫБРАТЬ {atr_name} КАК col_0 ИЗ Справочник.{cat_name} КАК {cat_name} ГДЕ {cat_name}.Наименование = "{naim}"'''
        query = self.V83.NewObject("Query", q)
        sel = query.Execute().Choose()
        sel = query.Execute().Select()  # -->  ВыборкаИзРезультатаЗапроса
        sel.next()
        return sel.Get(0)

    def getCatalogsVals(self, cat_ref):
        """Возвращает 1-ое значение реквизита записи справочника

        :param cat_ref: Ссылка на запись справочника
        :return: Значения реквизитов справочника. [ (key_1, val_1), (key_2, val_2), ...}
        """
        cat_name = cat_ref.Metadata().Name
        ls_attr = self.getCatalogsAtr(cat_name)
        naim = cat_ref.Наименование
        str_arg = ''.join([s + ' КАК col_' + str(i) + ', ' for i, s in enumerate(ls_attr)])[:-2]
        q = f'''ВЫБРАТЬ {str_arg} ИЗ Справочник.{cat_name} КАК {cat_name} ГДЕ {cat_name}.Наименование = "{naim}"'''
        ls = []
        query = self.V83.NewObject("Query", q)
        sel = query.Execute().Choose()
        sel = query.Execute().Select()  # -->  ВыборкаИзРезультатаЗапроса
        sel.next()
        [ls.append(self._classToVal(sel.Get(col[0]), col[1])) for col in enumerate(ls_attr)]
        return ls

    def _classToVal(self, obj, st):
        """Обработка полей в зависсимости от их класса

        :param obj: поле типа object
        :param st: наименование столбца
        :return:
        """
        if type(obj) == float:
            return (st, str(obj))
        elif type(obj) == int:
            return (st, str(obj))
        elif type(obj) == str:
            return (st, str(obj))
        elif type(obj) == win32com.client.CDispatch:
            try:
                return (st, obj.Наименование)
            except:
                (st, obj)
        # elif str(type(obj)) == 'pywintypes.datetime': return str(obj)
        else:
            return (st, str(obj))

    def loadFileToReg(self, doc_ref, img_src):
        """Загрузка бинарного файла в регистр сведений

        :return:
        """
        name_file = img_src.split('\\')[-1]
        bin = self.V83.NewObject("BinaryData", img_src)
        val_stor = self.V83.NewObject("ValueStorage", bin)  # --> ХранилищеЗначения
        # СоздатьМенеджерЗаписи
        reg_mng = getattr(self.V83.InformationRegisters,
                          "ХранилищеФайлов").CreateRecordManager()  # -->  РегистрСведенийМенеджерЗаписи
        reg_mng.Период = self._currDate()
        reg_mng.Объект = doc_ref
        reg_mng.Наименование = name_file  # 'Брэт_Пит'
        reg_mng.Файл = val_stor  # ХранилищеЗначения
        reg_mng.ИмяФайла = name_file
        reg_mng.ДатаФайла = self._currDate()
        reg_mng.МестоРазмещения = self.getLinkByName('Справочник.МестаРазмещенияЭлектроныхПриложений', 'В базе 1С')
        reg_mng.Write()

    def delFile(self):
        """Удаление записи регистра сведений "ХранилищеФайлов"

        :return:
        """
        ref_doc = self.docByNum('ЗаявкаНаВыполнениеРабот', number='000000460', data='20.01.2021  8:30:27')
        reg_mng = getattr(self.V83.InformationRegisters,
                          "ХранилищеФайлов").CreateRecordManager()  # -->  РегистрСведенийМенеджерЗаписи
        reg_mng.Период = '26.01.2021 5:55:48'
        reg_mng.Объект = ref_doc
        reg_mng.Наименование = "Брэт_Пит"
        reg_mng.Delete()


if __name__ == '__main__':
    conn_str = "Srvr=srv-01;Ref=rng;Usr='Иванов И.И';Pwd=123;"
    obj_1c_orl = Class_1C(conn_str)
    ls_atr_cat = obj_1c_orl.getCatalogsAtr('Банки')
    ls_atr_doc = obj_1c_orl.getDocAtr('АктСверки')
    print(ls_atr_cat)
    print(ls_atr_doc)
