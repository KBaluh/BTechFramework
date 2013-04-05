//-----------------------------------------------------------------------------
// BTechFramework - фреймворк для программирования в Terrasoft Administrator
//-----------------------------------------------------------------------------

var BTechFramework = { Version : 'v1.8.0', ModifyDate : '05.04.2013' };

//-----------------------------------------------------------------------------
// v1.8.0 - 05.04.2013
// ОБНОВЛЕНИЕ НЕ СОВМЕСТИМО С ВЕРСИЯМИ НИЖЕ 1.8.0
//
// Переименованы методы:
//	GetParamValueFromBPItem - GetBPValue
//	SetParamValueFromBPItem - SetBPValue
//	NotifyContainer - BTNotifyContainer
// Удалены методы:
// 	ExistDBFile
//
// Обновлены методы:
// SetBPValue - больше не возвращает булевский результат
// SetDatasetValue и SetDatasetValues - больше не возвращают булевский результат
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.7.3 - 28.03.2013
// Обновлены методы SetParamValueFromBPItem и GetParamValueFromBPItem,
// добавлены проверки на null
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.7.2 - 23.03.2013
// Обновлен метод метод SaveWithUpdate
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.7.1 - 20.02.2013
// Добавлен метод RemoveRightFull
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.7.0 - 19.02.2013
// Добавлен класс для работы с 1С синхронизацией
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.6.5 - 18.02.2013
// Добавлены комментарии к некоторым методам.
// Исправлен текст вывода сообщения ошибки в методе GetContainerDataset
// Исправлена потенциальная ошибка в Ajax при создании нового запроса
// Оптимизирован метод ChangeRightAccess, добавлен блок finally, что бы закрыть
// в конце датасет.
// Добавлен класс BTNotifyObject
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.6.1 - 15.02.2013
// Исправлена ошибка при добавлении прав доступа для записи, 
// не записывался AdminUnitID
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.6.0 - 14.12.2012
// Добавлен метод SaveWithUpdate.
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.5.0 - 13.12.2012
// Добавлен объект NotifyContainer, для хранения списка NotifyObject.
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.4.0 - 24.09.2012
// Добавлен метод ChangeRightAccess который изменят доступ к записи
// Добавлен метод по загрузке файла в базу, без создания экземпляра класса
// ExistDBFile - недоработанный метод по проверке файла в базе.
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.3.3 - 06.09.2012
// Права доступа - после удаления убран метод GotoNext
// Работа с датами - исправлена ошибка при создании объекта даты
// Убрано создание не используемого кэша для датасета
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.3.2 - 21.08.2012
// Добавлены методы по работе с курсором и потоком сообщения в блоке System
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.3.1 - 20.08.2012
// Добавлен метод OpenURL - фасад для выполнения Ajax запроса GET
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.3.0 - 15.08.2012
//
// Исправлено - Ajax
// - Убрано создание XMLHttpRequest, так как это поддержка только в браузера,
//   	он вызывал ошибку и переходил на создание Msxml2.XMLHTTP.
// - Добавлены константы REQUEST_STATUS_READY, REQUEST_STATUS_DONE отвечающие
// 		за готовность ответа в response объекте
//
// Доработано - BTechFramework
// - В BTechFramework объекте добавлена версия и дата изменения
//
// Добавлено - OpenOffice
// - Класс по работе с OpenOffice Writer: создание и открытие документов 
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.2.0 - 03.08.2012
//
// Добавлено - Ajax
// - Автор KARTOH
//
// - Доработан метод BTCatchException, добавлен параметр Message
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.1.0 - 26.07.2012
//
// Добавлено - HashMap, для хранения данных: ключ - значение
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.0.0 - Базовый функционал
//
// Работа с датасетом
// - Запись данных в одну запись
// - Чтение данных из одной записи
//
// Работа с окнами
// - Подключение детали в окно редактирования
// - Получение датасета из контейнера окна
//
// Работа с датами
// - Название месяцев на рус. и укр.
// - Время
//
// Работа с набором прав доступа
// - Добавление прав доступа и удаление
//
// Работа с почтовиком Thunderbird
// - Создание нового письма
// - Создание нового письма с заполнением
//
// Кэширование сервисов
// - Всех сервисов, датасетов, скриптов
//
// Исключения
// - Вывод сообщения ошибки с названием функции, в которой произошла ошибка
//
// UnitTest'ы
// - Тестирование кода
//
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// Dataset - функции для работы с датасетом
//
// Используемые скрипты: scr_DB
//-----------------------------------------------------------------------------

/**
 * Устанавливает в отфильтрированый датасет по ID, значение в поле.
 * SetDatasetValue('ds_Contact', ID_Контакта, 'Name', 'Значение которое запишется в поле Name');
 *
 * @param DatasetUSI - USI датасета
 * @param RecordID - ID записи
 * @param FieldName - Название поля, в которое запишется значение
 * @param Value - Значение
 */
function SetDatasetValue(DatasetUSI, RecordID, FieldName, Value) {
	var Parameters = new Object();
	Parameters[FieldName] = Value;	
	SetDatasetValues(DatasetUSI, RecordID, Parameters);
}

/**
 * Устанавливает в отфильтрированый датасет по ID, значения в поля.
 * SetDatasetValues('ds_Contact', ID_Контакта, {
 *	 'Name' : 'Значение которое запишется в поле Name',
 *	 'Description' : 'Значение которое запишется в поле Description'});
 * 
 * @param DatasetUSI - USI датасета
 * @param RecordID - ID записи
 * @param Parameters - объект параметров ключ-значение. Где ключ название поля.
 */
function SetDatasetValues(DatasetUSI, RecordID, Parameters) {
	var Dataset = DatasetCache_Get(DatasetUSI, 'ID');
	ApplyDatasetFilter(Dataset, 'ID', RecordID, true);
	Dataset.Open();
	if (IsDatasetEmpty(Dataset)) {
		BTCatchException('Не найдена запись в источнике данных: ' + 
			DatasetUSI + ', ID: ' + RecordID);
		return Result;
	}
	Dataset.Edit();
	for (var FieldName in Parameters) {
		Dataset(FieldName) = Parameters[FieldName];
	}
	try {
		Dataset.Post();
	} catch (Ex) {
		BTCatchException(Ex);
	}
}

/**
 * Возращает значение из датасета. Фильтрирует по ID
 * @param DatasetUSI - USI датасета
 * @param RecordID - ID записи
 * @param Field - название поля в датасете для получения значения
 * @return Value - значение
 */
function GetDatasetValue(DatasetUSI, RecordID, Field) {
	var Values = GetDatasetValues(DatasetUSI, RecordID, [Field]);
	var Value = Values[Field];
	return Value;	
}

/**
 * Возаращает значения из датасета. Фильтрирует по ID
 * @param DatasetUSI - USI датасета
 * @param RecordID - ID записи
 * @param Parameters - массив полей
 * @return Values - объект значений (поле, значение)
 */
function GetDatasetValues(DatasetUSI, RecordID, Parameters) {
	var Values = {};
	var Dataset = DatasetCache_Get(DatasetUSI, 'ID');
	ApplyDatasetFilter(Dataset, 'ID', RecordID, true);
	Dataset.Open();
	for (var i = 0; i < Parameters.length; ++i) {
		var FieldName = Parameters[i];
		try { 
			Values[FieldName] = Dataset(FieldName);
		} catch (Ex) {
			BTCatchException(Ex);
		}
	}
	Dataset.Close();
	return Values;
}

/**
 * Вызывает у датасета метод Post
 * Если перед сохранением было состояние на редактирование или добавление
 * тогда после сохранения он его вернет в состояние редактирования
 * @param Dataset - объект датасета который необходимо сохранить
 * @return Результат сохранения - true/false
 */
function SaveWithUpdate(Dataset) {
	if (Dataset == null) {
		BTCatchException("Argument excepted: Dataset");
		return false;
	}
	var State = Dataset.State;
	try {
		var ID = Dataset('ID');
		Dataset.Post();
		Dataset.Close();
		ApplyDatasetFilter(Dataset, 'ID', ID, true);
		Dataset.Open();
		return true;
	} catch (Ex) {
		BTCatchException(Ex);
		return false;
	} finally {
		if (State == dstEdit || State == dstInsert) {
			Dataset.Edit();
		}
	}
}

// ----------------------------------------------------------------------------
// Window - функции для работы с окнами
//
// Используемые скрипты: scr_DB
// ----------------------------------------------------------------------------

/**
 * Обновляет датасет контейнера
 * @param WindowContainer - окно контейнер в карточке редактирования
 */
function RefreshContainerDataset(WindowContainer) {
	if (!Assigned(WindowContainer)) {
		BTCatchException(
			'Невозможно обновить датасет контейнера, ' +
			'контейнер не определен');
		return;
	}
	var Dataset = GetContainerDataset(WindowContainer);
	RefreshDataset(Dataset);
}

/**
 * Получает из контейнера содержащего окно датасет
 * @return Dataset - может вернуть null
 */
function GetContainerDataset(WindowContainer) {
	var _NotAssignedContainer = 'не определен контейнер';
	var _NotAssignedWindow = 'не определено окно контейнера';
	
	var _PrintMessage = function(Message) {
		BTCatchException( 
			FormatStr('Невозможно вернуть датасет контейнера, %1', Message)); 
	};
	
	if (!Assigned(WindowContainer)) {
		_PrintMessage(_NotAssignedContainer);
		return null;
	}
	if (!Assigned(WindowContainer.Window)) {
		_PrintMessage(_NotAssignedWindow);
		return null;
	}
	
	var Dataset = WindowContainer.Window.ComponentsByName('dlData').Dataset;
	return Dataset;
}

/**
 * Подключение детали в окно редактирования
 * @param WindowContainer - окно контейнер с реестром, который необходимо подключить
 * @param DatasetUSI - датасет подключаемоего окна
 * @param FilterField - поле для фильтрации детали
 * @param RecordID - ID для фильтрации данных
 */  
function IncludeDetailEdit(WindowContainer, DatasetUSI, FilterField, RecordID) {
	var Script = GetScript('scr_WorkspaceUtils');
	Script.ScriptControl.Run('RefreshCommonDetail',
		null, WindowContainer, FilterField, FilterField, 
		DatasetUSI, null, null, null, null, RecordID);
}

/**
 * Копирование продуктов. Так же копирует и другие детали
 * @param SourceDataset - отфильтрованый датасет
 * @param SourField - поле по которому найдены продукты (OpportunityID, InvoiceID, ...)
 * @param DestinationDatasetUSI - USI датасета в который произойдет копирование
 * @param DestinationField - поле по которому будет найдены продукты
 * @param DestinationRecordID - значение по которому будут искаться продукты
 */
function CopyDetailData(SourceDataset, SourceField, 
		DestinationDatasetUSI, DestinationField, DestinationRecordID) {

	var DestinationDataset = Services.GetNewItemByUSI(DestinationDatasetUSI);

	// Данный параметр нужно установить в датасет для того,
	// что бы, не вытягивались подчиненные продукты.
	// Из за того что мы и так копируем с подчинением, то выходили дубликаты.	
	DestinationDataset.Attributes('IsTreeIgnore') = true;
	
	var Script = GetScript('scr_DB');
	Script.ScriptControl.Run('CopyTreeDetail',
		SourceDataset,
		DestinationDataset,
		SourceField,
		DestinationField,
		DestinationRecordID);
}

/**
 * Контейнер NotifyObject'ов. 
 * Добавить этот контейнер в атрибут NotifyObject
 * Когда произойдет вызов оповещения NotifyObject'а, 
 * он вызовет оповещение у всех своих объктов.
 */
function BTNotifyContainer() {

	/**
	 * Список всеъ объектов для оповещения
	 */
	var NotifyObjects = new Array();
	
	/**
	 * Проверяет, существует ли такой элемент уже в списке
	 * @param FindObject - искомый объект
	 * @return Индект найденого объекта, если такого нету вернет -1
	 */
	var IndexOf = function(FindObject) {
		for (var i = 0; i < NotifyObjects.length; ++i) {
			if (NotifyObjects[i] === FindObject) {
				return i;
			}
		}
		return -1;
	};
	
	/**
	 * Добавляет объект в список
	 * @param NotifyObject - объект с методом Notify(param1, param2, param3)
	 */
	this.Add = function(NotifyObject) {
	 	if (IndexOf(NotifyObject) < 0) {
			NotifyObjects.push(NotifyObject);	 
		}
	};
	
	/**
	 * Удаляет объект из контейнера
	 * @param NotifyObject - объкт для удаления
	 */
	this.Remove = function(NotifyObject) {
	 	var Index = IndexOf(NotifyObject);
	 	if (Index < 0) {
	 		return;
	 	}
	 	NotifyObjects.splice(Index, 1);
	};
	
	/**
	 * Метод который будет вызван оповещателем, вызывает у подписчиков метод Notify
	 * @param Sender - отправитель
	 * @param Message - сообщение
	 * @param Data - доп. данные
	 */
	this.Notify = function(Sender, Message, Data) {
		for (var i = 0; i < NotifyObjects.length; ++i) {
			var NotifyObject = NotifyObjects[i];
			if (Assigned(NotifyObject)) {
				NotifyObject.Notify(Sender, Message, Data);
			}
		}
	};
}

/**
 * NotifyObject - принимает оповещение. По умолчанию выводит сообщение в журнал сообщений.
 * Можно переопределить метод Notify, для этого необходимо присвоить новый:
 *
 * var NotifyObject = new BTNotifyObject();
 * NotifyObject.Notify = function(Sender, Message, Data) {
 * 		// TODO
 * };
 */
function BTNotifyObject() {
	
	/**
	 * Метод который будет вызван оповещателем
	 * @param Sender - отправитель
	 * @param Message - сообщение
	 * @param Data - доп. данные
	 */
	this.Notify = function(Sender, Message, Data) {
		Log.Write(1, '[BTNotifyObject] Message: ' + Message + ", Data: " + Data);
	};
}

//-----------------------------------------------------------------------------
// DateTime - работа с датами
//-----------------------------------------------------------------------------

/**
 * Возвращает серверную дату с временем
 * @return DateTime
 */
function GetServerDateTime() {
	return Connector.GetServerDateTime();
}

/**
 * Возвращает серверную дату
 */
function GetServerDate() {
	var DateTime = GetSystemDateTime();	
	var Result = ExtractDate(DateTime);
	return Result;
}

/**
 * Возвращает строку даты времени: день.месяц.год часы:минуты:секунды
 * @param DateValue
 * @return TimeStr {String} - 10.05.2012 09:00:03
 */
function GetDateTimeStr(DateValue) {
	var DateObject = new Date(DateValue);
	var DateStr = DateToStr(DateObject);
	var TimeStr = "00:00:00";
	
	var Hours = DateObject.getHours();
	var Minutes = DateObject.getMinutes();
	var Seconds = DateObject.getSeconds();
	
	if (Hours < 10) Hours = "0" + Hours;
	if (Minutes < 10) Minutes = "0" + Minutes;
	if (Seconds < 10) Seconds = "0" + Seconds;
	
	TimeStr = DateStr + " " + Hours + ":" + Minutes + ":" + Seconds; 
	return TimeStr;
}

/**
 * 10.05.2012 00:00:00, 10.05.2012 23:59:59 
 * @param DateValue
 * @return Arr {Array} [0] - начало дня, [1] конец дня
 */
function GetDateBeginEnd(DateValue) {
	var SplitStr = ".";
	
	var DateObject = new Date(DateValue);
	
	var day = DateObject.getDate();
	var month = DateObject.getMonth() + 1;
	var year = DateObject.getFullYear();
	if (day < 10) {
		day = "0" + day; 
	}
	if (month < 10) {
		month = "0" + month; 
	}
	
	var Begin = day + SplitStr + month + SplitStr + year + " " +
		"00:00:00";
	var End = day + SplitStr + month + SplitStr + year + " " +
		"23:59:59";
	var Arr = [Begin, End];
	return Arr;
}

/**
 * Возвращает  название месяца.
 * @param MonthIndex - номер месяца (нумерация с нуля)
 * @param Multipli - Множество
 * @param Launge - язык на котором вернет название месяца (RU, UA)
 *		по умолчанию используется русский.
 * @return Name {String} - название месяца
 */ 
function GetMonthName(MonthIndex, Launge, Multipli) {
	var MonthRU1 = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль",
  		"Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"];
  	var MonthRU2 = ["Января", "Февраля", "Марта", "Апреля", "Мая", "Июня", "Июля",
  		"Августа", "Сентября", "Октября", "Ноября", "Декабря"];
	
	var MonthUA1 = ["Січень", "Лютий", "Березень", "Квітень", "Травень", "Червень", 
		"Липень", "Серпень", "Вересень", "Жовтень", "Листопад", "Груднень"];
	var MonthUA2 = ["Січня", "Лютня", "Березня", "Квітня", "Травня", "Червня", 
		"Липня", "Серпня", "Вересня", "Жовтня", "Листопада", "Грудня"];

	if (!Boolean(Launge)) {
		Launge = 'RU';
	}
	
	var MonthLength = 11;
	if (MonthIndex > MonthLength) {
		MonthIndex = MonthIndex - MonthLength - 1;
	}
	
	var Name = '';
	switch(Launge) {
		case 'UA':
			Name = (Multipli ? MonthUA2[MonthIndex] : MonthUA1[MonthIndex]);
			break;
		case 'RU':
		default:
			Name = (Multipli ? MonthRU2[MonthIndex] : MonthRU1[MonthIndex]); 
			break;	
	}
	return Name;
}

/**
 * @test - получения названия месяца
 */
function testGetMonthName() {
	Assert.startTest("GetMonthName");

	// Русский
	var Month = GetMonthName(0, 'RU');
	Assert.equals("Январь", Month);
	
	Month = GetMonthName(11, 'RU');
	Assert.equals("Декабрь", Month);

	Month = GetMonthName(13, 'RU');
	Assert.equals("Февраль", Month);

	Month = GetMonthName(0, 'RU', true);
	Assert.equals("Января", Month);
	
	// Украинский
	Month = GetMonthName(0, 'UA');
	Assert.equals("Січень", Month);
	
	Month = GetMonthName(11, 'UA');
	Assert.equals("Груднень", Month);
	
	Month = GetMonthName(13, 'UA');
	Assert.equals("Лютий", Month);

	Month = GetMonthName(0, 'UA', true);
	Assert.equals("Січня", Month);
	
	Assert.finishTest();
}

//-----------------------------------------------------------------------------
// Workflow - методы по работе с бизнес процессами
//
// Используемые скрипты: scr_WorkflowUtils
//-----------------------------------------------------------------------------

/**
 * Запускает БП, пример параметров:
 * Parameters = 
 * @param WorkflowUSI - USI диаграммы БП
 * @param Parameters - Объект, содержащий свойства как поля со значениями:
 * 		{ "ContactID" : {123-123-123-123}, "AccountID" : {321-321-321-321} }
 */
function RunWorkflow(WorkflowUSI, Parameters) {
	var Script = GetScript('scr_WorkflowUtils');
	var WorkflowEngine = GetWorkflowEngine();	
	var Now = new Date().getVarDate();
	var Params = System.CreateObject('TSWorkflowLibrary.WorkflowParameters');
	for (Parameter in Parameters) {
		Script.ScriptControl.Run('WFSetParamValueDirect', 
			Params, Parameter, Parameters[Parameter]);	
	}
	WorkflowEngine.StartWorkflow(WorkflowUSI, Now, Params);
}

/**
 * Возвращает workflow engine. В разной версии он разный, поэтому нужно
 * настроить метод
 * @return WorkflowEngine
 */
function GetWorkflowEngine() {
	var WorkflowEngineAttrName = 'WorkflowEngine';
	var WorkflowEngine = Connector.Attributes(WorkflowEngineAttrName);
	if (!WorkflowEngine) {
		WorkflowEngine = System.CreateObject('TSWorkflowLibrary.WorkflowEngine');
		Connector.Attributes(WorkflowEngineAttrName) = WorkflowEngine;
		WorkflowEngine.Connector = Connector;
	}
	return WorkflowEngine;
}

/**
 * Возвращает значение параметра из элемента БП
 * @param Item - элемент процесса
 * @param ParameterName - имя параметра в диаграмме
 * @return Value - значение
 */
function GetBPValue(Item, ParameterName) {
	if (!Assigned(Item)) {
		BTCatchException("[GetBPValue].[Item] is null");
		return null;
	}
	var ParentItems = Item.ParentItems;
	if (!Assigned(ParentItems)) {
		BTCatchException("[GetBPValue].[Item.ParentItems] is null");
		return null;
	}
	
	var ParentDiagram = ParentItems.ParentDiagram;
	if (!Assigned(ParentDiagram)) {
		BTCatchException("[GetBPValue].[Item.ParentItems.ParentDiagram] is null");
		return null;
	}
	
	var Parameters = ParentDiagram.Parameters;
	if (!Assigned(Parameters)) {
		BTCatchException("[GetBPValue].[Item.ParentItems.ParentDiagram.Parameters] is null");
		return null;
	}
	
	var Value = null;
	try {
		Value = Parameters.ItemsByName(ParameterName).Value;
	} catch (Ex) {
		BTCatchException("Ошибка получения параметра диаграмы: " + ParameterName);
	}
	return Value;
}

/**
 * Устанавливает в параметр значение
 * @param Item - элемент процесса
 * @param ParameterName - имя параметра в диаграмму
 * @param Value - значение
 */
function SetBPValue(Item, ParameterName, Value) {
	if (!Assigned(Item)) {
		BTCatchException("[SetBPValue].[Item] is null");
		return null;
	}
	var ParentItems = Item.ParentItems;
	if (!Assigned(ParentItems)) {
		BTCatchException("[SetBPValue].[Item.ParentItems] is null");
		return null;
	}
	
	var ParentDiagram = ParentItems.ParentDiagram;
	if (!Assigned(ParentDiagram)) {
		BTCatchException("[SetBPValue].[Item.ParentItems.ParentDiagram] is null");
		return null;
	}
	
	var Parameters = ParentDiagram.Parameters;
	if (!Assigned(Parameters)) {
		BTCatchException("[SetBPValue].[Item.ParentItems.ParentDiagram.Parameters] is null");
		return null;
	}
	
	try {
		Parameters.ItemsByName(ParameterName).Value = Value;
	} catch (e) {
		BTCatchException('Ошибка изменения значения параметру диаграмы: ' + 
			ParameterName + ', значение: ' + Value);
	}
}

//-----------------------------------------------------------------------------
// Right - права доступа
//-----------------------------------------------------------------------------

/**
 * Изменяет права доступа для записи
 * @param TableUSI - таблица прав (пример: tbl_ContactRight
 * @param RecordID - запись к которой подчиняются доступы
 * @AdminUnitID - ID записи из датасета ds_AdminUnit
 * @param CanRead - доступ на чтение
 * @param CanWrite - доступ на редактирование
 * @param CanChange - доступ на изменение доступа
 */
function ChangeRightAccess(TableUSI, RecordID, AdminUnitID, 
	CanRead, CanWrite, CanDelete, CanChange) {
	
	var Script = GetScript('scr_Access');	 
	var Dataset = Script.ScriptControl.Run('GetItemRightDataset', TableUSI);
	ApplyDatasetFilter(Dataset, 'AdminUnitID', AdminUnitID, true);
	ApplyDatasetFilter(Dataset, 'RecordID', RecordID, true);
	Dataset.Open();
	try {
		if (IsDatasetEmpty(Dataset)) {
			return;	
		}	
		Dataset.Edit();
		Dataset('CanRead') = Boolean(CanRead);
		Dataset('CanWrite') = Boolean(CanWrite);
		Dataset('CanDelete') = Boolean(CanDelete);
		Dataset('CanChangeAccess') = Boolean(CanChange);
		Dataset.Post();
	} catch (Ex) {
		BTCatchException(Ex);
	} finally {
		Dataset.Close();
	}
}

/**
 * Удаление группы/пользователя из доступов для всех записей
 * @param TableUSI - таблица прав (пример: tbl_ContactRight)
 * @AdminUnitID - ID записи из датасета ds_AdminUnit 
 */ 
function RemoveRightAll(TableUSI, AdminUnitID) {
	var Script = GetScript('scr_Access');	 
	var Dataset = Script.ScriptControl.Run('GetItemRightDataset', TableUSI);
	ApplyDatasetFilter(Dataset, 'AdminUnitID', AdminUnitID, true);
	Dataset.FetchRecordsCount = -1;
	Dataset.Open();
	while(!Dataset.IsEOF) {
		Dataset.Delete();
	}		
	Dataset.Close();
}

/**
 * Удаление всех прав доступов для всех записей таблицы
 * @param TableUSI - таблица прав (пример: tbl_ContactRight)
 */ 
function RemoveRightFull(TableUSI) {
	var Script = GetScript('scr_Access');	 
	var Dataset = Script.ScriptControl.Run('GetItemRightDataset', TableUSI);
	Dataset.FetchRecordsCount = -1;
	Dataset.Open();
	while(!Dataset.IsEOF) {
		Dataset.Delete();
	}		
	Dataset.Close();
}

/**
 * Удаление группы/пользователя из доступов для одной записи
 * @param RecordID - запись, к которой подчиняются доступы
 * @param TableUSI - таблица прав (пример: tbl_ContactRight)
 * @AdminUnitID - ID записи из датасета ds_AdminUnit 
 */ 
function RemoveRightRecord(TableUSI, RecordID, AdminUnitID) {
	var Script = GetScript('scr_Access');
	var Dataset = Script.ScriptControl.Run('GetItemRightDataset', TableUSI);
	ApplyDatasetFilter(Dataset, 'AdminUnitID', AdminUnitID, true);
	ApplyDatasetFilter(Dataset, 'RecordID', RecordID, true);
	Dataset.Open();
	while(!Dataset.IsEOF) {
		try {
			Dataset.Delete();
		} catch (Ex) {
			BTCatchException(Ex);
		}
	}
	Dataset.Close();
}

/**
 * Добавление прав доступа для одной записи
 * Если добавляемая группа существует, будут установлены новые
 * @param TableUSI - таблица прав (пример: tbl_ContactRight)
 * @AdminUnitID - ID записи из датасета ds_AdminUnit
 * @param CanRead - доступ на чтение
 * @param CanWrite - доступ на редактирование
 * @param CanChange - доступ на изменение доступа
 * @param IsGroup - это группа(true) или контакт(false)
 */ 
function AddRightRecord(RecordID, TableUSI, AdminUnitID, 
		CanRead, CanWrite, CanDelete, CanChange, IsGroup) {
	
	var IsExistRight = function(RightDataset, AdminUnitID) {
		var Result = false;
		RightDataset.GotoFirst(); 
		while(!RightDataset.IsEOF) {
			if (RightDataset('AdminUnitID') == AdminUnitID) {
				Result = true;
				break;	
			}
			RightDataset.GotoNext();
		}
		RightDataset.Close();
		return Result;
	};
	
	var Script = GetScript('scr_Access');
	var RightDataset = Script.ScriptControl.Run('GetItemRightDataset', TableUSI);
	RightDataset.FetchRecordsCount = -1;
	ApplyDatasetFilter(RightDataset, 'RecordID', RecordID, true); 
	RightDataset.Open();	
	if (IsExistRight(RightDataset, AdminUnitID)) {
		RightDataset.Edit();
		RightDataset('CanRead') = CanRead;
		RightDataset('CanWrite') = CanWrite;
		RightDataset('CanDelete') = CanDelete;
		RightDataset('CanChangeAccess') = CanChange;
		RightDataset('IsGroup') = IsGroup;
		try {
			RightDataset.Post();
		} catch (Ex) {
			BTCatchException(Ex);
		}
	} else {
		RightDataset.Append();
		RightDataset('RecordID') = RecordID;
		RightDataset('AdminUnitID') = AdminUnitID;
		RightDataset('CanRead') = CanRead;
		RightDataset('CanWrite') = CanWrite;
		RightDataset('CanDelete') = CanDelete;
		RightDataset('CanChangeAccess') = CanChange;
		RightDataset('IsGroup') = IsGroup; 
		try {
			RightDataset.Post();
		} catch (Ex) {
			BTCatchException(Ex);
		}
	}
	RightDataset.Close();
}

/**
 * Добавление прав доступа всем записям из датасета.
 * Если добавляемая группа существует, будут установлены новые
 * @param DatasetUSI - датасет которому будет добавлена группа
 * @param TableUSI - таблица прав (пример: tbl_ContactRight)
 * @AdminUnitID - ID записи из датасета ds_AdminUnit
 * @param CanRead - доступ на чтение
 * @param CanWrite - доступ на редактирование
 * @param CanChange - доступ на изменение доступа
 * @param IsGroup - это группа(true) или контакт(false)
 */
function AddRightAll(DatasetUSI, TableUSI, AdminUnitID, CanRead, CanWrite, 
		CanDelete, CanChange) {
		
	var Dataset = Services.GetNewItemByUSI(DatasetUSI);
	Dataset.Open();
	while(!Dataset.IsEOF) {
		var RecordID = Dataset('ID');
		
		AddRightRecord(RecordID, TableUSI, AdminUnitID, 
			CanRead, CanWrite, CanDelete, CanChange);
					
		Dataset.GotoNext();
	}
	Dataset.Close();
}

//-----------------------------------------------------------------------------
// Files - работа с файлами
//
// Используемые скрипты: scr_FileUtils
//-----------------------------------------------------------------------------

function Files() {

	var _InsertQueryUSI = '';
	var _SubjectFieldName = '';
	
	/**
	 * @param Dataset - датасет файлов
	 * @param ObjectType - тип файла (ссылка, файл, путь к папке)
	 * @param OldFileID
	 * @param ShortFileName - название файла для отображения в детали
	 * @param LongFileName - полный путь к файлу
	 * @param Revision - поставить 1
	 * @param AddedFileIDsArray - просто массив
	 * @param SubjectID - ID записи к которой добавляется файл
	 * @param IsLog - вывести сообщение ID файла
	 */
	var _LoadToFile = function(Dataset, ObjectType, OldFileID, ShortFileName, 
		LongFileName, Revision, AddedFileIDsArray, SubjectID, IsLog) {
		var ID = Connector.GenGUID();
		try {
			Dataset.Append();
			Dataset.Values('ID') = ID;
			Dataset.Values('ItemTypeID') = ObjectType;
			Dataset.Values('Link') = ShortFileName;
			Dataset.Values('Revision') = Revision;
			if (ObjectType == ObjectType) {
				var Script = GetScript('scr_FileUtils');
				var Result = Script.ScriptControl.Run('SaveFileToDataset', 
					LongFileName, Dataset, 'FileData');
				if (!Result) {
					return ;
				}
			}
			var FilePosted = Dataset.Post();
		} finally {
			Dataset.Close();
		}
		 if (FilePosted) {
			_LinkFile(ID, SubjectID);
		}
	};
	
	/**
	 * @param FileID - ссылка на файл из таблицы файлов
	 * @param SubjectID
	 */
	var _LinkFile = function(FileID, SubjectID) {
		try {
			var InsertQuery = Services.GetNewItemByUSI(_InsertQueryUSI);
			var ColumnsValues = InsertQuery.ColumnsValues;
			ColumnsValues.ItemsByName('FileID').Value = FileID;
			ColumnsValues.ItemsByName(_SubjectFieldName).Value = SubjectID;
			var ID = Connector.GenGUID();
			ColumnsValues.ItemsByName('ID').Value = ID;
			InsertQuery.Execute();
		} catch(e) { }
	};
	
	this._LoadDBFile = function(FilePatch, DisplayName, 
		InsertQueryUSI, SubjectID, SubjectFieldName) {
		
		_InsertQueryUSI = InsertQueryUSI;
		_SubjectFieldName = SubjectFieldName;
		
		var FilesDataset = Services.GetNewItemByUSI('ds_Files');
		// Тип загружаемого файла - Файл.
		var ObjectType = "{39A5B367-4A7A-473E-8F74-26977CB6DB67}";
		var ShortFileName = DisplayName; 
		var LongFileName = FilePatch;
		var Revision = 1;
		var AddedFileIDsArray = new Array();
		_LoadToFile(FilesDataset, ObjectType, null, ShortFileName, 
		LongFileName, Revision, AddedFileIDsArray, SubjectID);
	};
}

Files.prototype = {
	LoadDBFile: function(FilePatch, DisplayName, 
		InsertQueryUSI, SubjectID, SubjectFieldName)
	{
		this._LoadDBFile(FilePatch, DisplayName, InsertQueryUSI, 
			SubjectID, SubjectFieldName);
	},
	
	/**
	 * Извлекает имя файла из его пути
	 * @param FilePatch - путь к файлу
	 */
	ExtractFileName : function(FilePatch) {
		var FilePatchArr = FilePatch.split('\\');
		var FileName = FilePatchArr[FilePatchArr.length - 1];
		return FileName;
	}
};

/**
 * Загрузка файла в деталь "Файлы"
 * @param FilePatch - путь к файлу на компьютере
 * @param DisplayName - отображаемое имя
 * @param InsertQueryUSI - сервис добавления
 * @param SubjectID - ID записи к которой будет произведено добавление файла
 * 					  например к контакту или контрагенту
 * @param SubjectFieldName - название поля в котором хранится SubjectID
 */
function LoadDBFile(FilePatch, InsertQueryUSI, SubjectID, SubjectFieldName) {
	var DisplayName = DBFiles.ExtractFileName(FilePatch);
	
	DBFiles.LoadDBFile(FilePatch, DisplayName, InsertQueryUSI, 
		SubjectID, SubjectFieldName);	
}

//-----------------------------------------------------------------------------
// Exceptions - исключения ошибок
//-----------------------------------------------------------------------------

/**
 * Exception метод, отображает в каком методе произошла ошибка и его сообщение
 * @param Exception - Exception object
 * @param Message - произвольное сообще об ошибке
 */
function BTCatchException(Ex, Message) {
	var ParseFunctionName = function(Caller) {
		var Pattern = /function (.*[^\n\r\{ ])/;
		var Result = Pattern.exec(Caller);
		var FunctionName = "";
		if (Result != null) {
			FunctionName = Result[1].toString();
		}
		return FunctionName;
	};
	var Caller = ParseFunctionName(arguments.callee.caller);
	
	var ErrMessage = "";
	if (Ex instanceof Error) {
		ErrMessage = Ex.message;
	} else
	if ((Ex instanceof String) || (typeof(Ex) == 'string')) {
		ErrMessage = Ex;
	} else
	if (Ex.hasOwnProperty('message')) {
		ErrMessage = Ex.message;
	} else
	if (Ex.hasOwnProperty('Message')) {
		ErrMessage = Ex.Message;
	} else
	if (Ex instanceof Object) {
		ErrMessage = Ex.toString();
	} else {
		ErrMessage = "Неизвестная ошибка: " + Ex.toString();
	}
	
	var MessageStr = (Message ? 
		FormatStr('Произошла ошибка в методе %1: %2, оригинальное сообщение - %3', 
			Caller, Message, ErrMessage) :
		FormatStr('Произошла ошибка в методе %1: %2', 
			Caller, ErrMessage));
	Log.Write(2, MessageStr);
}

// ----------------------------------------------------------------------------
// Caching system - кеширование объектов
// ----------------------------------------------------------------------------

/**
 * Система кэширования. Содержит в себе кэшируемые объекты
 * Используется патерн Singleton для одного экземпляра объекта
 * Объекты бывают трех видов: скрипты, датасеты, остальные сервисы
 */
var CacheSystem = new function() {
	var instance;
		
	// Кэш для сервисов
	var CacheItem = {};
	
	// Кэш для датасетов
	var CacheDataset = {};
	
	// Кэш для скриптов
	var CacheScript = {};
	
	function CacheSystem() {
		if ( !instance )
			instance = this;
		else return instance;
	}

	CacheSystem.prototype.GetItem = function(ItemUSI, Key) {
		var KeyFiled = ItemUSI + Key;
		var Item = CacheItem[KeyFiled];
		if (Item == null) {
			Item = Services.GetNewItemByUSI(ItemUSI);
			CacheItem[KeyFiled] = Item;
			LastItemKey = KeyFiled;
			return Item;
		} else {
			LastItemKey = KeyFiled;
			return Item;
		}	
	};
	
	CacheSystem.prototype.GetDataset = function(ItemUSI, Key) {
		var KeyFiled = ItemUSI + Key;
		var Item = CacheDataset[KeyFiled];
		if (Item == null) {
			Item = Services.GetNewItemByUSI(ItemUSI);
			CacheDataset[KeyFiled] = Item;
			return Item;
		} else {
			if (Item != null && Item.ServiceTypeCode != null) {
				if (Item.ServiceTypeCode == 'DBDataset') {
					Item.Close();
				}
			}
			return Item;
		}	
	};
	
	CacheSystem.prototype.GetScript = function(ItemUSI) {
		var Item = CacheScript[ItemUSI];
		if (Item == null) {
			Item = Services.GetNewItemByUSI(ItemUSI);
			CacheScript[ItemUSI] = Item;
			return Item;
		} else {
			return Item;
		}	
	};
	
	return CacheSystem;
}

/**
 * Возвращает из кэша датасет.
 * Если не находит - создает новый датасет и возвращает.
 * @param DatasetUSI - название датасета
 * @param Key - ключ уникальности
 * @return Датасет
 */
function DatasetCache_Get(DatasetUSI, Key) {	
	return CS.GetDataset(DatasetUSI, Key);		
}

/**
 * Возвращает скрипт и кэширует его в системе.
 * @param ScriptUSI - USI скрипта
 * @return Скрипт
 */
function GetScript(ScriptUSI) {
	return CS.GetScript(ScriptUSI);		
}

/**
 * Кэширует и возвращает элемент
 * @ItemUSI - сервис в системе
 * @Key - ключ уникальности
 * @return Объект
 */
function CacheItem(ItemUSI, Key) {
	return CS.GetItem(ItemUSI, Key);
}

//-----------------------------------------------------------------------------
// Thunderbird - отправка писем через почтовик Thunderbird
//
// Пример использования:
//
// Создать новое письмо
// var TH = new Thunderbird();
// TH.CreateNewEmail();
//
// Создать новое с заполнением письмо
// var TH = new Thunderbird();
// TH.SendSingleEmail("support@btech.com", "Моя тема", "<b>Мой текст сообщения</b>");
//-----------------------------------------------------------------------------

function Thunderbird() { }

Thunderbird.prototype = {
	constructor: Thunderbird,
		
	/**
	 * Создаем пустое сообщение
	 */
	CreateNewEmail: function()
	{
		var WshShell = new ActiveXObject("WScript.Shell");
		WshShell.Run('thunderbird -compose');
	},
	
	/**
	 *	To - Эмайл кому отправлять (MyEmail@mail.com)
	 *	Subject - Тема сообщения (Dear friend!)
	 *	Body - Текст сообщения (Dear friend, i'm fine....)
	 *	AttachmentList - массив, который содержит ссылки на файлы в компьютере
	 */
	SendSingleEmail: function(To, Subject, Body, AttachmentList)
	{
		var WshShell = new ActiveXObject("WScript.Shell");
		try {
		WshShell.Run("thunderbird -compose \"to='" + To + "\',subject='" + 
			Subject + "',body='" + Body + "'");
		} catch(e) {
			Log.Write(2, "Произошла ошибка при отправке, проверьте установлен ли почтовик. " + e.message);
		} finally {
			WshShell = System.EmptyValue;
		}
	}
};

//-----------------------------------------------------------------------------
// UnitTest - тестирование скриптов
//
// Пример использования
// Assert.equals(5, 2 + 3);
//-----------------------------------------------------------------------------

/**
 * Unit test
 *
 * -> equals
 *	if values equals, test passed
 *	  
 * -> isTrue
 *	if condition is true, test passed
 *
 * -> isFalse
 *	if condition is false, test passed
 *
 * -> isFail
 * 	if calback functin throw exception, test passed
 *
 * -> isNotFail
 *	if calback function dont throw exception, test passed
 *
 * -> parameter [MessageOff] 
 *	if the parameter is set to "true" then do not get a message
 *	Example:	
 *		Assert.isFalse(Assert.isTrue(4 == 5, true));
 *		You get a passed message.
 *           
 * @author Baluh Konstantin <baluh.kostya@gmail.com>
 */
function UnitTest() {
	this.Name = "UnitTest";
}

UnitTest.prototype = {
	__info: function(Message) {
		Log.Write(1, this.Name + ": " + Message);
	},
	
	__error: function(Message) {
		Log.Write(2, this.Name + ": " + Message);
	},
	
	startTest: function(TestName) {
		this.Name = (TestName || "UnitTest");
		this.__info("Test started");
	},
	
	finishTest: function() {
		this.__info("Finish test");
		this.Name = "UnitTest";	
	},
	
	equals: function(Object1, Object2, MessageOff) {
		var r = false;	
		if (Object1 === Object2) {
			if (!MessageOff)
				this.__info("\tPassed");
			r = true;
		} else {
			if (!MessageOff)
				this.__error("\tFail");
		}
		return r;
	},
	
	isTrue: function(Condition, MessageOff) {
		var r = false;
		if (Condition) {
			if (!MessageOff)
				this.__info("\tPassed");
			r = true;
		} else {
			if (!MessageOff)
				this.__error("\tFail");
		}
		return r;
	},
	
	isFalse: function(Condition, MessageOff) {
		var r = false;
		if (Condition) {
			if (!MessageOff)
				this.__error("\tFail");
		} else {
			if (!MessageOff) 
				this.__info("\tPassed");
			r = true;
		}
		return r;
	},
	
	// TODO - rewrite args, this is Array
	isFail: function(calback, args, MessageOff) {
		var r = false;
		try {
			calback(args);
			if (!MessageOff)
				this.__error("\tFail");
		} catch (e) {
			if (!MessageOff)
				this.__info("\tPassed");
			r = true;
		}
		return r;
	},
	
	// TODO - rewrite args, this is Array
	isNotFail: function(calback, args, MessageOff) {
	    var r = false;
		try {
			calback(args);
			if (!MessageOff)
				this.__info("\tPassed");
			r = true;
		} catch (e) {
			if (!MessageOff)
				this.__error("\tFail");
		}
		return r;
	}
};

//-----------------------------------------------------------------------------
// HashMap
//-----------------------------------------------------------------------------

function HashMap() {
	
	var Map = {};
	
	this.Add = function(Key, Value) {
		Map[Key] = Value;	
	};
	
	this.Get = function(Key) {
		var Value = Map[Key];
		return Value;
	};
	
	this.Clear = function() {
		Map = {};
		CollectGarbage();
	};
	
	this.GetMap = function() {
		return Map;
	};
}

//-----------------------------------------------------------------------------
// Ajax
//-----------------------------------------------------------------------------
/**
 * Класс для роботы с Ajax
 * @author KARTOH
 * @author KBaluh {baluh.kostya@gmail.com}
 * @date 03.08.2012
 * @version 1.2
 */
var Ajax = new function() {
	
	/**
	 * Инстанс класса
	 */
	var instance; 
	
	var REQUEST_STATUS_READY = 4;
	var REQUEST_STATUS_DONE = 200;
 
	/***
	Конструктор - используется патерн Singleton
	*/
	function Ajax() {
		if (!instance)
			instance = this;
		else return instance;
	}

	/**
	 * Хранилище Запросов {Array}
	 */
	this.requests = [];
	
	/**
	 * Метод поиска свободного запроса.
	 * Если свободного запроса нету, создает новый
	 */
	this.get_request = function() {
		if (!this.requests.length) {
			this.requests[0] = this.create_request();
			return this.requests[0];
		}
		
		for (var i = 0; i < this.requests.length; i++) {
			if (this.requests[i].readyState == REQUEST_STATUS_READY) {
				return this.requests[i];
			}
		}
		
		var indexNewRequest = ++i;
		var request = this.create_request();
		this.requests[indexNewRequest] = request;
		
		return request;
	};
	
	/**
	 * Метод создания запроса
	 */
	this.create_request = function() {
		var request;
	
		try {
			// Получение объекта IE >= 6 версии
			request = new ActiveXObject("Msxml2.XMLHTTP");
		}
		catch (Ex) {
			try {
				// Получение объекта IE <= 5 версии
				request = new ActiveXObject("Microsoft.XMLHTTP");
			}
			catch (ExIE_5) {
				BTCatchException(ExIE_5,
					"[Ajax] Невозможно создать XMLHttpRequest");
				request = false;
			}
		}
		return request;
	};
	
	/**
	 * Метод отправки запроса
	 * @param {String} url - Адрес отправки
	 * @param {String} send - Отправляемые параметры
	 * @param {Function} callback - Отправляет в обратную функцию текст страницы
	 * @param {Object} head - Заголовки
	 * @param {Boolean} asunc - Синхроное / Асинхроное выполнение запросов
	 * 							По умолчанию синхронный
	 */
	this.send_request = function(url, send, callback, head, async) {
		var request = this.get_request();
		if (!request) {
			return false;
		}
		
		var http = send ? 'POST' : 'GET';
		request.open(http, url, Boolean(async));
		
		// Присваивание заголовков
		for(var key in head) {
			request.setRequestHeader(key,head[key]);
		}
		
		// Отправка запроса
		request.onreadystatechange = function () {
			if (request.readyState == REQUEST_STATUS_READY) {
				if (request.status == REQUEST_STATUS_DONE) {
					callback(request.responseText);	
				} else {
					BTCatchException('[Ajax] Error connect. Status Error #' + 
						request.status + ' Error text = ' + 
						request.getResponseHeader("Content-Type"));
				}
			}
		};
		
		try {
			request.send(send ? send : null);
		} catch (Ex) {
			BTCatchException(Ex, '[Ajax] Error on send request');
		}
	};
	
	return this;
};

/**
 * Фасад для GET запроса на открытие страницы
 * Открывает ссылку и возвращает текст страницы
 * @param Url
 * @return Result {String} - текст страницы
 */
function OpenURL(Url) {
	var Result = null;
	
	var Callback = function(ResponseText) {
		Result = ResponseText;	
	};
	
	Ajax.send_request(Url, false, Callback);  
	
	return Result;
}

/**
 * @test - проверяет отправку запроса с "null" head'ером
 */
function testSendRequestNullHead() {

	/**
	 * Анонимный метод, выводит ответ от отправки сервера
	 */
	var testAjaxCallbackMethod = function(ResponseText) {
		Log.Write(1, 'Получен ответ сервера: ' + ResponseText);
	};

	Ajax.send_request("http://translate.google.com.ua/#en/ru/callback", false,
		testAjaxCallbackMethod, null, null); 		
}

//-----------------------------------------------------------------------------
// Open Office
//-----------------------------------------------------------------------------

function OpenOffice() {

	var _ServiceManager = null;
	var _Desktop = null;

	var _Documents = new Array();
	var _ActiveDocument = null;
	var _DocumentIndex = 0;
	
	/**
	 * Выдает в консоль сообще об ошибке, что не установлен ОО
	 */
	var ThrowInstallError = function() {
		BTCatchException("Не удалось получить OpenOffice. " +
			"Проверьте правильность установки и установленных компонентов ActiveXObject.", 
			"[OpenOffice.ThrowInstallError]");				
	};
	
	/**
	 * Устанавливает в текущий документ - документ
	 * @param Document
	 */
	var SetActiveDocument = function(Document) {
		_Documents.push(Document);
		_ActiveDocument = Document;
	};

	/**
	 * Возвращает объект по работе с ОО
	 * @return Desktop
	 */
	var GetDesktop = function() {
		if (_Desktop != null) {
			return _Desktop;
		}	
		try {
			if (_ServiceManager == null) {
				_ServiceManager = new ActiveXObject("com.sun.star.ServiceManager");
			}
			try { 
				_Desktop = _ServiceManager.createInstance("com.sun.star.frame.Desktop");
				return _Desktop;
			} catch (Ex) {
				BTCatchException(Ex);
				return null;
			}  
		} catch (Ex) {
			return null;
		}
	};

	/**
	 * Открывает файл
	 * @param Patch
	 */	
	this.OpenDocument = function(Patch) {
		var Desktop = GetDesktop();
		if (Desktop == null) {
			ThrowInstallError();	
			return;
		}
		
		// a new document 
		var args = new Array(); 
		var Url = "file:///" + Patch;
		var Document = Desktop.loadComponentFromURL(Url, "_blank", 0, args);
		
		SetActiveDocument(Document);	
	};
	
	/**
	 * Создает новый текстовый документ
	 * @return objDocument
	 */
	this.CreateNewWriter = function() {
		var Desktop = GetDesktop();
		if (Desktop == null) {
			ThrowInstallError();
			return;
		}
			
		var args = new Array(); 
		var Document = Desktop.loadComponentFromURL("private:factory/swriter",
			"_blank", 0, args);
		
		SetActiveDocument(Document);
	};
	
	/**
	 * Возвращает активный документ
	 * @return _ActiveDocument
	 */
	this.GetActiveDocument = function() {
	 	return _ActiveDocument;
	};
	
	this.HasNext = function() {
		if (_Documents.length > 0) {
			if (_Documents.length < _DocumentIndex + 1) {
				return true;	
			} else {
				return false;
			}
		} else {
			return false;
		}
	};
	
	this.NextDocument = function() {
		_DocumentIndex++;
		var Document = _Documents[_DocumentIndex];
		return Document;	
	};
	
	this.First = function() {
		_DocumentIndex = 0;	
	};
}

//-----------------------------------------------------------------------------
// System
//-----------------------------------------------------------------------------

/**
 * Делает прерывание для вызова лог сообщения
 */
function MessageThread() {
	System.ProcessMessages();
}

/**
 * Ставит внешний вид курсора "Ожидание"
 */
function CursorWait() {
	System.BeginProcessing();	
}

/**
 * Ставит внешний вид курсора "По умолчанию"
 */
function CursorDefault() {
	System.EndProcessing();	
}

// ----------------------------------------------------------------------------
// 1C - функции для работы с 1C
// ----------------------------------------------------------------------------

/**
 * Запускает элемент интеграции
 * @param DataflowID - ID интеграции
 * @param ItemID - ID элемента в интеграции
 */
function RunOneCSyncItem(DataflowID, ItemID) {
	var Script = GetScript('scr_Dataflow1CUtils');
	Script.ScriptControl.Run('ExportObject', null, DataflowID, ItemID);  
}

/**
 * Класс для подключения к базе 1С
 */
function ConnectOneC() {
	var _user = "";
	var _password = "";
	var _path = "";

	var _app = null;
	var _connected = false; 

	this.ConnectFromDataflow = function(DataflowID) {
		var Obj = new Object();
		var Script = GetScript('scr_Dataflow1CUtils');
		var Param = Script.ScriptControl.Run('LoadSettings', DataflowID, Obj);
		var UserName = Param.UserName;
		var Password = Param.Password;
		var Path = Param.Path;
		return this.Connect(UserName, Password, Path);   
	};

	this.Connect = function(User, Password, Path) {
		_user = User;
		_password = Password;
		_path = Path;

		try {
			if (_app == null) {
				_app = new ActiveXObject("V82.Application");
			}
		} catch (Ex) {
			BTCatchException(Ex);
			return false;
		}  
		try {
			_connected = _app.Connect("File=\"" + _path +  
			"\";Usr=\"" + _user + "\";Pwd=\"" + _password + "\";"); 
		} catch (Ex) {
			_connected = false;
			BTCatchException(Ex);
		}
		return _connected;
	};

	this.Disconnect = function() {
		if (Assigned(_app)) {
			_app.Exit(false);
			return true;
		} else {
			return false;
		}
	};

	this.GetApp = function() {
		return _app;
	};

	this.Connected = function() {
		return _connected;
	};
}

// ----------------------------------------------------------------------------
// Mock objects - объекты с интерфейсом нормальных объектов
// ----------------------------------------------------------------------------

/**
 * Датасет
 */
function MockDataset() {

	var dstInactive = 0x00000000;
	var dstBrowse = 0x00000001;
	var dstEdit = 0x00000002;
	var dstInsert = 0x00000003;
	var dstCalcFields = 0x00000004;
	
	var DataValues = {};

	this.State = dstInactive;	
	
	this.Open = function() {
	 	this.State = dstBrowse;
	};
	
	this.Close = function() {
	 	this.State = dstInactive; 
	};
	
	this.Append = function() {
	 	this.State = dstInsert;
	};
	
	this.Edit = function() {
		this.State = dstEdit;	
	};
	
	this.Post = function() {
		this.State = dstBrowse;
	};
	
	this.Values = function(FieldName) {
		return DataValues[FieldName];
	};
	
	this.AppendField = function(FieldName) {
		if (DataValues[FieldName] == undefined) {
			DataValues[FieldName] = null;
		}
	};
	
	this.RemoveField = function(FieldName) {
		DataValues[FieldName] = null;
		delete DataValues[FieldName];
	};
}

/**
 * Окно
 */
function MockWindow() {
	this.Attributes	= GetNewDictionary();
	this.Attributes.Add('DefaultValues', GetNewDictionary());	
}

/**
 * Параметры БП
 */
function MockParameters() {
	var Params = new Array();

	this.Add = function(Item) {
	 	Params.push(Item);
	};
	
	this.ItemsByName = function(ItemName) {
		var SearchItem = null;
		for (var i = 0; i < Params.length; ++i) {
			var Param = Params[i];
			if (Param.Name == ItemName) {
				SearchItem = Param;
				break;	
			} 
		}
		return SearchItem;
	};	
}

/**
 * Параметры карты БП
 */
function MockParametersMap() {
	var Params = new MockParameters();

	this.Add = function(Item) {
	 	Params.Add(Item);
	};
	
	this.ItemsByItemParameterName = function(ItemName) {
		return Params.ItemsByName(ItemName);
	};	
}

/**
 * Параметр карты БП
 */
function MockParameterMap() {
	this.Name = "";
	this.Value = null;	
}

/**
 * Элемент БП
 */
function MockActionItem() {
	this.ParametersMap = new MockParametersMap();
	this.Parameters = new MockParameters();
	this.ParentItems = new MockParameters();
	this.ParentItems.ParentDiagram = new MockDiagram();	
}

/**
 * Диаграмма БП
 */
function MockDiagram() {
	this.Parameters = new MockParameters();
}

//-----------------------------------------------------------------------------
// Init framework
//-----------------------------------------------------------------------------

// Кеширование
var CS = new CacheSystem();

// Создание юнит тестов
var Assert = new UnitTest();

// Файлы
var DBFiles = new Files();

//-----------------------------------------------------------------------------
// Private methods - если не подключать scr_DB, этих методов достаточно для
// работы данного скрипта
//-----------------------------------------------------------------------------

function ApplyDatasetFilter(Dataset, FilterName, ParamValue, Enabled) {
	var Script = GetScript('scr_DB');
	Script.ScriptControl.Run('ApplyDatasetFilter',
		Dataset, FilterName, ParamValue, Enabled);
}

function IsDatasetEmpty(Dataset) {
	var Script = GetScript('scr_DB');
	return Script.ScriptControl.Run('IsDatasetEmpty', Dataset);
}

function Assigned(obj) {
    return ((obj != null) && (typeof(obj) == ObjectTypeName))
}

function ShowConfirmationDialog(Message) {
    return System.MessageDialog(Message, mdtConfirmation, mdbYes + mdbNo, 0);
} 

//-----------------------------------------------------------------------------
// Run
//-----------------------------------------------------------------------------

function Main() {

}
