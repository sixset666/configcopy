#Если Клиент Тогда
////////////////////////////////////////////////////////////////////////////////
// ПЕРЕМЕННЫЕ ОБЪЕКТА

Перем мКуда Экспорт;			// Содержит назначение
Перем мОтправитель Экспорт;		// Отправитель
Перем мПолучатель Экспорт;		// Получатель

////////////////////////////////////////////////////////////////////////////////
// ЭКСПОРТНЫЕ ПРОЦЕДУРЫ И ФУНКЦИИ ДОПОЛНИТЕЛЬНЫХ МЕТОДОВ ОБЪЕКТА

// Функция определения адреса объекта
//
Функция ОпределитьАдресОбъекта(Объект) Экспорт
	
	Если НЕ ЗначениеЗаполнено (Объект) Тогда
		Возврат "";	
	КонецЕсли; 
	
	Запрос = Новый Запрос;
	
	ТекстЗапроса = "ВЫБРАТЬ ПЕРВЫЕ 1
	               |	КонтактнаяИнформация.Представление
	               |ИЗ
	               |	РегистрСведений.КонтактнаяИнформация КАК КонтактнаяИнформация
	               |ГДЕ
	               |	КонтактнаяИнформация.Объект = &Объект
	               |	И КонтактнаяИнформация.Тип = &Тип
	               |	И КонтактнаяИнформация.Вид = &Вид";
				   
	Запрос.Текст = ТекстЗапроса;
	Запрос.УстановитьПараметр("Объект",Объект);
	Запрос.УстановитьПараметр("Тип",Перечисления.ТипыКонтактнойИнформации.Адрес);
	
	Если ТипЗнч(Объект) = Тип("СправочникСсылка.Контрагенты") Тогда
		ВидАдреса = ВидАдресаКонтрагента;
	Иначе
		ВидАдреса = ВидАдресаОрганизации;
	КонецЕсли;
	
	Запрос.УстановитьПараметр("Вид",ВидАдреса);
		
	ВыборкаАдресов = Запрос.Выполнить().Выбрать();
	
	Пока ВыборкаАдресов.Следующий() Цикл
		
		Наименование = Объект.НаименованиеПолное; 

		Возврат ВыборкаАдресов.Представление + Символы.ПС + Наименование+Символы.ПС;
		
	КонецЦикла;		
	
КонецФункции	

// Функция получения списка видов адресов контрагента
//
Функция ПолучитьСписокВидовАдресовКонтрагента(Объект="") Экспорт
	
	//сформируем список видов адресов контрагента
	Запрос = Новый Запрос;
	Запрос.Текст = "ВЫБРАТЬ
	|	ВидыКонтактнойИнформации.Ссылка
	|ИЗ
	|	Справочник.ВидыКонтактнойИнформации КАК ВидыКонтактнойИнформации
	|ГДЕ
	|	ВидыКонтактнойИнформации.Тип = &Тип";
	
	Запрос.УстановитьПараметр("Тип",Перечисления.ТипыКонтактнойИнформации.Адрес);
	
	Результат = Запрос.Выполнить();
	Выборка = Результат.Выбрать();
	
	СписокВидовАдресов = Новый СписокЗначений;
	
	Пока Выборка.Следующий() Цикл
	
		СписокВидовАдресов.Добавить(Выборка.Ссылка,Выборка.Ссылка.Наименование);
	
	КонецЦикла;
	
	Возврат СписокВидовАдресов; 

КонецФункции	

// Работа с адресом получателя
//
Функция АдресПолучателя(Контрагент) Экспорт
	
	АдресПолучателя = ОпределитьАдресОбъекта(Контрагент);
	
	Если НЕ ЗначениеЗаполнено (АдресПолучателя) И Не ПолучательПоОбразцу Тогда
		Сообщить ("Не указан адрес контрагента <"+Контрагент+">");
		Возврат Неопределено
	Иначе	
		Возврат мКуда + АдресПолучателя + Символы.ПС+ спПолучитьПредставление(Контрагент);
	КонецЕсли;	
		
КонецФункции	

// Процедура печати конвертов в word
//
Процедура ПечатьКонвертовWORD () Экспорт
	
	Если  Контрагенты.Количество() = 0 Тогда 
		Возврат;
	КонецЕсли;      
	
	
	ИмяФайлаАдресов = КаталогВременныхФайлов()+"address.dbf";
	Текст = Новый XBase();
	Текст.Поля.Добавить("Address","S",254);
	Текст.Поля.Добавить("RAddress","S",254);
	Текст.Кодировка = ?(КодировкаANSI,КодировкаXBase.ANSI,КодировкаXBase.OEM);
	Текст.СоздатьФайл(ИмяФайлаАдресов);
	Если Текст.Открыта() Тогда
		Текст.ЗакрытьФайл();
	КонецЕсли;
	Текст.ОткрытьФайл(ИмяФайлаАдресов);
	Если НЕ Текст.Открыта() Тогда
		Возврат;
	КонецЕсли;      
	
	КоличествоКонвертов = 0;
	
	Для каждого Строка Из  Контрагенты  Цикл
		
		Если НЕ Строка.Выбран Тогда
			Продолжить
		КонецЕсли;      
		
		АдресПолучателя =  АдресПолучателя(Строка.Контрагент);
		
		Если АдресПолучателя = Неопределено Тогда
			Продолжить;
		Конецесли;      
		
		Текст.Добавить();
		Текст.Address = АдресПолучателя;
		Текст.RAddress = АдресОтправителя;
		Текст.Записать();
		КоличествоКонвертов = КоличествоКонвертов+1;
		
	КонецЦикла;
	Текст.ЗакрытьФайл();
	
	Если КоличествоКонвертов = 0 Тогда 
		Возврат; 
	КонецЕсли; 
	
	Попытка
		Word = Новый ComОбъект("Word.Application");
	Исключение
		Предупреждение(ОписаниеОшибки() + Символы.ПС + "программа Word не установлена на данном компьютере!");
		Возврат;
	КонецПопытки;
	
	Попытка
		Состояние("Формирование документа Word...");
		
		Word.Documents.Add(,,,Истина);  //Add(Template, NewTemplate, DocumentType, Visible)
		
		//Template   Optional Variant. The name of the template to be used for the new document. If this argument is omitted, the Normal template is used.
		
		//NewTemplate   Optional Variant. True to open the document as a template. The default value is False.
		
		//DocumentType   Optional Variant. Can be one of the following WdNewDocumentType constants: wdNewBlankDocument, wdNewEmailMessage, wdNewFrameset, or wdNewWebPage. The default constant is wdNewBlankDocument.
		
		//Visible   Optional Variant. True to open the document in a visible window. If this value is False, Microsoft Word opens the document but sets the Visible property of the document window to False. The default value is True.
		
		//Envelope  - конверт
		
		//ExtractAddress   Optional Variant. True to use the text marked by the EnvelopeAddress bookmark (a user-defined bookmark) as the recipient's address.
		
		//Address   Optional Variant. A string that specifies the recipient's address (ignored if ExtractAddress is True).
		
		//AutoText   Optional Variant. A string that specifies an AutoText entry to use for the address. If specified, Address is ignored.
		
		//OmitReturnAddress   Optional Variant. True to not insert a return address.
		
		//ReturnAddress   Optional Variant. A string that specifies the return address.
		
		//ReturnAutoText   Optional Variant. A string that specifies an AutoText entry to use for the return address. If specified, ReturnAddress is ignored.
		
		//PrintBarCode   Optional Variant. True to add a POSTNET bar code. For U.S. mail only.
		
		//PrintFIMA   Optional Variant. True to add a Facing Identification Mark (FIMA) for use in presorting courtesy reply mail. For U.S. mail only.
		
		
		Если РазмерыКонверта = "Нестандартный размер" Тогда
			Word.ActiveDocument.Envelope.DefaultHeight = Word.CentimetersToPoints (Высота);
			Word.ActiveDocument.Envelope.DefaultWidth = Word.CentimetersToPoints (Ширина);
		Иначе 
			Размеры = СформировтьРазмерыКонверта();
			Word.ActiveDocument.Envelope.DefaultHeight = Word.CentimetersToPoints (Размеры.Высота);
			Word.ActiveDocument.Envelope.DefaultWidth = Word.CentimetersToPoints (Размеры.Ширина);
		КонецЕсли;
		
		Word.ActiveWindow.WindowState = 1; // состояние окна 0 обычное, 1 свернуто, 2 развернуто
		
		Попытка
			Word.ActiveDocument.Envelope.Insert (Истина,"","", НЕ ПечататьАдресОтправителя,"","",Ложь,Ложь);//ExtractAddress, Address, AutoText, OmitReturnAddress, ReturnAddress, ReturnAutoText, PrintBarCode, PrintFIMA
		Исключение
			Сообщить(ОписаниеОшибки());
		КонецПопытки;
		
		Word.ActiveDocument.Envelope.AddressFromLeft = 0;// "wdAutoPosition"; 
		Word.ActiveDocument.Envelope.AddressFromTop = 0;// "wdAutoPosition"; 
		Word.ActiveDocument.Envelope.ReturnAddressFromLeft = 0;//"wdAutoPosition"; 
		Word.ActiveDocument.Envelope.ReturnAddressFromTop = 0;//"wdAutoPosition"; 
		
		
		Word.ActiveDocument.Envelope.DefaultFaceUp = Истина;  //True to print the envelope face up, False to print it face down.
		//Word.ActiveDocument.Envelope.DefaultOrientation = ; //Optional Variant. The orientation for the envelope.
		//#define wdLeftPortrait  0
		//#define wdCenterPortrait  1
		//#define wdRightPortrait  2
		//#define wdLeftLandscape  3
		//#define wdCenterLandscape  4
		//#define wdRightLandscape  5
		//#define wdLeftClockwise  6
		//#define wdCenterClockwise  7
		//#define wdRightClockwise  8
		
		
		//Попытка
		//      Word.ActiveDocument.Envelope.Vertical = Ложь; //True to print vertical text on the envelope. Used for Asian envelopes. Default is False.
		//Исключение
		//КонецПопытки;
		
		
		doc = Word.ActiveDocument;
		
		Word.ActiveDocument.MailMerge.MainDocumentType = "wdEnvelopes";
		Word.ActiveDocument.MailMerge.OpenDataSource (ИмяФайлаАдресов,"wdOpenFormatAuto", Ложь, Истина);
		Word.ActiveDocument.MailMerge.Fields.Add (Word.ActiveDocument.Envelope.Address,"Address");
		Если ПечататьАдресОтправителя  Тогда
			Word.ActiveDocument.MailMerge.Fields.Add (Word.ActiveDocument.Envelope.ReturnAddress,"RAddress");
		КонецЕсли;      
		Word.ActiveDocument.MailMerge.Destination = "wdSendToNewDocument";
		Word.ActiveDocument.MailMerge.SuppressBlankLines = Ложь;
		Word.ActiveDocument.MailMerge.DataSource.FirstRecord = Истина;
		Word.ActiveDocument.MailMerge.DataSource.LastRecord = КоличествоКонвертов;
		Word.ActiveDocument.MailMerge.Execute (Истина);                                                                                         
		
		LastDocument = Word.ActiveDocument;
		doc.Close(Ложь);
		
		//+ Удаление пустых листов
		Если НЕ НЕУдалятьАвтоматическиПустыеЛистыWord Тогда
			i = 1;
			p = 1;
			КоличествоСимволов = LastDocument.Characters.Count;
			Пока LastDocument.Characters.Item (i) <> LastDocument.Characters.Last Цикл
				Если КодСимвола(LastDocument.Characters.Item (i).Text) = 12 Тогда
					Если p - Цел(p/2)*2 = 0 Тогда
						LastDocument.Characters.Item (i).Delete();
					КонецЕсли;      
					p = p + 1;
				КонецЕсли;      
				Если i =  LastDocument.Characters.Count - 2 Тогда
					Прервать;
				КонецЕсли;      
				i = i + 1;
			КонецЦикла;
		КонецЕсли;
		
		//-
		//+Форматирование текста
		myRange = LastDocument.Range();
		Пока myRange.Find.Execute(СокрЛП(мКуда)) <> 0 Цикл
			myRange.Bold = Истина;
		КонецЦикла;                        
		myRange = LastDocument.Range();
		Пока myRange.Find.Execute(СокрЛП(мОтправитель)) <> 0 Цикл
			myRange.Bold = Истина;                
		КонецЦикла;                        
		myRange = LastDocument.Range();
		Пока myRange.Find.Execute(СокрЛП(мПолучатель)) <> 0 Цикл
			myRange.Bold = Истина;                
		КонецЦикла;                 
		//-Форматирование текста
	Исключение
		Сообщить(ОписаниеОшибки());
		Word.Quit();
		Возврат;
	КонецПопытки;
	
	Word.Visible = 1;
КонецПроцедуры

// Процедура печати наклеек mxl
//
Процедура ПечатьНаклеекMXL () Экспорт
	
	Если  Контрагенты.Количество() = 0 Тогда 
		Возврат;
	КонецЕсли;	
	
	ТабличныйДокумент  = Новый ТабличныйДокумент; 
	
	Если СокрЛП(ПутьКШаблону) = ""  Тогда
		Макет = ПолучитьМакет("ШаблонНаклеек");
	Иначе
		
		Попытка 
			Макет = Новый ТабличныйДокумент; 
			Макет.Прочитать(ПутьКШаблону);
		Исключение
			Сообщить(ОписаниеОшибки());
		КонецПопытки;
	КонецЕсли;
	
	ОбластьМакета = Макет.ПолучитьОбласть("Горизонтальная|Вертикальная");	

	Столбцов = 2;
	Сч = 0;
	Для каждого Строка Из  Контрагенты Цикл
		Если   АдресДляПечатиНаклеек = 0 Тогда
			Адрес =  АдресПолучателя(Строка.Контрагент);
			Если  Адрес = Неопределено Тогда Продолжить КонецЕсли;;
		Иначе
			Адрес = АдресОтправителя;
		КонецЕсли;	
		
		ОбластьМакета.Параметры.Адрес = Адрес;
		Если Сч < Столбцов Тогда
			ТабличныйДокумент.Присоединить(ОбластьМакета);
			Сч = Сч+1;
			Продолжить;
		КонецЕсли;	
		ТабличныйДокумент.Вывести(ОбластьМакета);	
		Сч = 1;
	КонецЦикла;	
	
	Наименование = ?(АдресДляПечатиНаклеек = 0,"(Адреса получателей)","(Адреса отправителей)"); 
	
	ТабличныйДокумент.Показать("Почтовые наклейки" + Наименование);
	
КонецПроцедуры

// Процедура печати наклеек word
//
Процедура ПечатьНаклеекWORD () Экспорт
	
	Если  Контрагенты.Количество() = 0 Тогда 
		Возврат;
	КонецЕсли;	
	
	Состояние("Формирование документа Word...");	
	
	Попытка
		Word = Новый ComОбъект("Word.Application");
	  	Исключение
	   	Предупреждение(ОписаниеОшибки() + Символы.ПС + "программа Word не установлена на данном компьютере!");
	   	Возврат;
	КонецПопытки;
	Word.Documents.Add("",0,,1);
	Word.Application.MailingLabel.DefaultPrintBarCode = Ложь;
	Word.ActiveWindow.ActivePane.VerticalPercentScrolled = Ложь;
	
	Для каждого Строка Из  Контрагенты  Цикл
		Если   АдресДляПечатиНаклеек = 0 Тогда
			Адрес =  АдресПолучателя(Строка.Контрагент);
			Если  Адрес = Неопределено Тогда Продолжить КонецЕсли;;
		Иначе
			Адрес = АдресОтправителя;
		КонецЕсли;
		Word.Application.MailingLabel.CreateNewDocument (,Адрес);
	КонецЦикла;	
	Word.Visible = 1;	
	
КонецПроцедуры

// Функция формирования размера конверта
//
Функция  СформировтьРазмерыКонверта() Экспорт
	
	Размеры = Новый Структура;
	
	Если РазмерыКонверта = "C6" Тогда
		Размеры.Вставить("Ширина",16.2);
		Размеры.Вставить("Высота",11.4);
	ИначеЕсли РазмерыКонверта = "E65" Тогда 
	    Размеры.Вставить("Ширина",22.0);
		Размеры.Вставить("Высота",11.0);
	ИначеЕсли РазмерыКонверта = "C5" Тогда 
	    Размеры.Вставить("Ширина",22.9); 	
		Размеры.Вставить("Высота",16.2);
	ИначеЕсли РазмерыКонверта = "C4" Тогда 
	    Размеры.Вставить("Ширина",32.4); 
		Размеры.Вставить("Высота",22.9);
	ИначеЕсли РазмерыКонверта = "B4" Тогда 
	    Размеры.Вставить("Ширина",35.3);
		Размеры.Вставить("Высота",25.0);
	ИначеЕсли РазмерыКонверта = "K7" Тогда 
		Размеры.Вставить("Ширина",14.0);
		Размеры.Вставить("Высота",9.0);
	ИначеЕсли РазмерыКонверта = "K6" Тогда 
		Размеры.Вставить("Ширина",12.5);
		Размеры.Вставить("Высота",12.5);
	ИначеЕсли РазмерыКонверта = "K65" Тогда 
		Размеры.Вставить("Ширина",19.0);
		Размеры.Вставить("Высота",12.5);
	ИначеЕсли РазмерыКонверта = "K5" Тогда 
		Размеры.Вставить("Ширина",21.5);
		Размеры.Вставить("Высота",14.5);
	ИначеЕсли РазмерыКонверта = "K8" Тогда 
		Размеры.Вставить("Ширина",15.0);
		Размеры.Вставить("Высота",15.0);
	КонецЕсли;	
	
	Возврат Размеры;
	
КонецФункции	

////////////////////////////////////////////////////////////////////////////////
// ПРОЦЕДУРЫ - ОБРАБОТЧИКИ СТАНДАРТНЫХ СОБЫТИЙ ОБЪЕКТА


////////////////////////////////////////////////////////////////////////////////
// ИСПОЛНЯЕМАЯ ЧАСТЬ МОДУЛЯ

мКуда = "Куда: ";
мОтправитель = "Отправитель: ";
мПолучатель = "Получатель: ";

РазмерыКонверта = "Нестандартный размер";
Высота = 11;
Ширина = 22;

#КонецЕсли