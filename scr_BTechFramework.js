//-----------------------------------------------------------------------------
// BTechFramework - ��������� ��� ���������������� � Terrasoft Administrator
//-----------------------------------------------------------------------------

var BTechFramework = { Version : 'v1.8.0', ModifyDate : '05.04.2013' };

//-----------------------------------------------------------------------------
// v1.8.0 - 05.04.2013
// ���������� �� ���������� � �������� ���� 1.8.0
//
// ������������� ������:
//	GetParamValueFromBPItem - GetBPValue
//	SetParamValueFromBPItem - SetBPValue
//	NotifyContainer - BTNotifyContainer
// ������� ������:
// 	ExistDBFile
//
// ��������� ������:
// SetBPValue - ������ �� ���������� ��������� ���������
// SetDatasetValue � SetDatasetValues - ������ �� ���������� ��������� ���������
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.7.3 - 28.03.2013
// ��������� ������ SetParamValueFromBPItem � GetParamValueFromBPItem,
// ��������� �������� �� null
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.7.2 - 23.03.2013
// �������� ����� ����� SaveWithUpdate
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.7.1 - 20.02.2013
// �������� ����� RemoveRightFull
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.7.0 - 19.02.2013
// �������� ����� ��� ������ � 1� ��������������
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.6.5 - 18.02.2013
// ��������� ����������� � ��������� �������.
// ��������� ����� ������ ��������� ������ � ������ GetContainerDataset
// ���������� ������������� ������ � Ajax ��� �������� ������ �������
// ������������� ����� ChangeRightAccess, �������� ���� finally, ��� �� �������
// � ����� �������.
// �������� ����� BTNotifyObject
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.6.1 - 15.02.2013
// ���������� ������ ��� ���������� ���� ������� ��� ������, 
// �� ����������� AdminUnitID
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.6.0 - 14.12.2012
// �������� ����� SaveWithUpdate.
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.5.0 - 13.12.2012
// �������� ������ NotifyContainer, ��� �������� ������ NotifyObject.
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.4.0 - 24.09.2012
// �������� ����� ChangeRightAccess ������� ������� ������ � ������
// �������� ����� �� �������� ����� � ����, ��� �������� ���������� ������
// ExistDBFile - �������������� ����� �� �������� ����� � ����.
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.3.3 - 06.09.2012
// ����� ������� - ����� �������� ����� ����� GotoNext
// ������ � ������ - ���������� ������ ��� �������� ������� ����
// ������ �������� �� ������������� ���� ��� ��������
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.3.2 - 21.08.2012
// ��������� ������ �� ������ � �������� � ������� ��������� � ����� System
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.3.1 - 20.08.2012
// �������� ����� OpenURL - ����� ��� ���������� Ajax ������� GET
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.3.0 - 15.08.2012
//
// ���������� - Ajax
// - ������ �������� XMLHttpRequest, ��� ��� ��� ��������� ������ � ��������,
//   	�� ������� ������ � ��������� �� �������� Msxml2.XMLHTTP.
// - ��������� ��������� REQUEST_STATUS_READY, REQUEST_STATUS_DONE ����������
// 		�� ���������� ������ � response �������
//
// ���������� - BTechFramework
// - � BTechFramework ������� ��������� ������ � ���� ���������
//
// ��������� - OpenOffice
// - ����� �� ������ � OpenOffice Writer: �������� � �������� ���������� 
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.2.0 - 03.08.2012
//
// ��������� - Ajax
// - ����� KARTOH
//
// - ��������� ����� BTCatchException, �������� �������� Message
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.1.0 - 26.07.2012
//
// ��������� - HashMap, ��� �������� ������: ���� - ��������
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// v1.0.0 - ������� ����������
//
// ������ � ���������
// - ������ ������ � ���� ������
// - ������ ������ �� ����� ������
//
// ������ � ������
// - ����������� ������ � ���� ��������������
// - ��������� �������� �� ���������� ����
//
// ������ � ������
// - �������� ������� �� ���. � ���.
// - �����
//
// ������ � ������� ���� �������
// - ���������� ���� ������� � ��������
//
// ������ � ���������� Thunderbird
// - �������� ������ ������
// - �������� ������ ������ � �����������
//
// ����������� ��������
// - ���� ��������, ���������, ��������
//
// ����������
// - ����� ��������� ������ � ��������� �������, � ������� ��������� ������
//
// UnitTest'�
// - ������������ ����
//
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// Dataset - ������� ��� ������ � ���������
//
// ������������ �������: scr_DB
//-----------------------------------------------------------------------------

/**
 * ������������� � ���������������� ������� �� ID, �������� � ����.
 * SetDatasetValue('ds_Contact', ID_��������, 'Name', '�������� ������� ��������� � ���� Name');
 *
 * @param DatasetUSI - USI ��������
 * @param RecordID - ID ������
 * @param FieldName - �������� ����, � ������� ��������� ��������
 * @param Value - ��������
 */
function SetDatasetValue(DatasetUSI, RecordID, FieldName, Value) {
	var Parameters = new Object();
	Parameters[FieldName] = Value;	
	SetDatasetValues(DatasetUSI, RecordID, Parameters);
}

/**
 * ������������� � ���������������� ������� �� ID, �������� � ����.
 * SetDatasetValues('ds_Contact', ID_��������, {
 *	 'Name' : '�������� ������� ��������� � ���� Name',
 *	 'Description' : '�������� ������� ��������� � ���� Description'});
 * 
 * @param DatasetUSI - USI ��������
 * @param RecordID - ID ������
 * @param Parameters - ������ ���������� ����-��������. ��� ���� �������� ����.
 */
function SetDatasetValues(DatasetUSI, RecordID, Parameters) {
	var Dataset = DatasetCache_Get(DatasetUSI, 'ID');
	ApplyDatasetFilter(Dataset, 'ID', RecordID, true);
	Dataset.Open();
	if (IsDatasetEmpty(Dataset)) {
		BTCatchException('�� ������� ������ � ��������� ������: ' + 
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
 * ��������� �������� �� ��������. ����������� �� ID
 * @param DatasetUSI - USI ��������
 * @param RecordID - ID ������
 * @param Field - �������� ���� � �������� ��� ��������� ��������
 * @return Value - ��������
 */
function GetDatasetValue(DatasetUSI, RecordID, Field) {
	var Values = GetDatasetValues(DatasetUSI, RecordID, [Field]);
	var Value = Values[Field];
	return Value;	
}

/**
 * ���������� �������� �� ��������. ����������� �� ID
 * @param DatasetUSI - USI ��������
 * @param RecordID - ID ������
 * @param Parameters - ������ �����
 * @return Values - ������ �������� (����, ��������)
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
 * �������� � �������� ����� Post
 * ���� ����� ����������� ���� ��������� �� �������������� ��� ����������
 * ����� ����� ���������� �� ��� ������ � ��������� ��������������
 * @param Dataset - ������ �������� ������� ���������� ���������
 * @return ��������� ���������� - true/false
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
// Window - ������� ��� ������ � ������
//
// ������������ �������: scr_DB
// ----------------------------------------------------------------------------

/**
 * ��������� ������� ����������
 * @param WindowContainer - ���� ��������� � �������� ��������������
 */
function RefreshContainerDataset(WindowContainer) {
	if (!Assigned(WindowContainer)) {
		BTCatchException(
			'���������� �������� ������� ����������, ' +
			'��������� �� ���������');
		return;
	}
	var Dataset = GetContainerDataset(WindowContainer);
	RefreshDataset(Dataset);
}

/**
 * �������� �� ���������� ����������� ���� �������
 * @return Dataset - ����� ������� null
 */
function GetContainerDataset(WindowContainer) {
	var _NotAssignedContainer = '�� ��������� ���������';
	var _NotAssignedWindow = '�� ���������� ���� ����������';
	
	var _PrintMessage = function(Message) {
		BTCatchException( 
			FormatStr('���������� ������� ������� ����������, %1', Message)); 
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
 * ����������� ������ � ���� ��������������
 * @param WindowContainer - ���� ��������� � ��������, ������� ���������� ����������
 * @param DatasetUSI - ������� �������������� ����
 * @param FilterField - ���� ��� ���������� ������
 * @param RecordID - ID ��� ���������� ������
 */  
function IncludeDetailEdit(WindowContainer, DatasetUSI, FilterField, RecordID) {
	var Script = GetScript('scr_WorkspaceUtils');
	Script.ScriptControl.Run('RefreshCommonDetail',
		null, WindowContainer, FilterField, FilterField, 
		DatasetUSI, null, null, null, null, RecordID);
}

/**
 * ����������� ���������. ��� �� �������� � ������ ������
 * @param SourceDataset - �������������� �������
 * @param SourField - ���� �� �������� ������� �������� (OpportunityID, InvoiceID, ...)
 * @param DestinationDatasetUSI - USI �������� � ������� ���������� �����������
 * @param DestinationField - ���� �� �������� ����� ������� ��������
 * @param DestinationRecordID - �������� �� �������� ����� �������� ��������
 */
function CopyDetailData(SourceDataset, SourceField, 
		DestinationDatasetUSI, DestinationField, DestinationRecordID) {

	var DestinationDataset = Services.GetNewItemByUSI(DestinationDatasetUSI);

	// ������ �������� ����� ���������� � ������� ��� ����,
	// ��� ��, �� ������������ ����������� ��������.
	// �� �� ���� ��� �� � ��� �������� � �����������, �� �������� ���������.	
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
 * ��������� NotifyObject'��. 
 * �������� ���� ��������� � ������� NotifyObject
 * ����� ���������� ����� ���������� NotifyObject'�, 
 * �� ������� ���������� � ���� ����� �������.
 */
function BTNotifyContainer() {

	/**
	 * ������ ���� �������� ��� ����������
	 */
	var NotifyObjects = new Array();
	
	/**
	 * ���������, ���������� �� ����� ������� ��� � ������
	 * @param FindObject - ������� ������
	 * @return ������ ��������� �������, ���� ������ ���� ������ -1
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
	 * ��������� ������ � ������
	 * @param NotifyObject - ������ � ������� Notify(param1, param2, param3)
	 */
	this.Add = function(NotifyObject) {
	 	if (IndexOf(NotifyObject) < 0) {
			NotifyObjects.push(NotifyObject);	 
		}
	};
	
	/**
	 * ������� ������ �� ����������
	 * @param NotifyObject - ����� ��� ��������
	 */
	this.Remove = function(NotifyObject) {
	 	var Index = IndexOf(NotifyObject);
	 	if (Index < 0) {
	 		return;
	 	}
	 	NotifyObjects.splice(Index, 1);
	};
	
	/**
	 * ����� ������� ����� ������ ������������, �������� � ����������� ����� Notify
	 * @param Sender - �����������
	 * @param Message - ���������
	 * @param Data - ���. ������
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
 * NotifyObject - ��������� ����������. �� ��������� ������� ��������� � ������ ���������.
 * ����� �������������� ����� Notify, ��� ����� ���������� ��������� �����:
 *
 * var NotifyObject = new BTNotifyObject();
 * NotifyObject.Notify = function(Sender, Message, Data) {
 * 		// TODO
 * };
 */
function BTNotifyObject() {
	
	/**
	 * ����� ������� ����� ������ ������������
	 * @param Sender - �����������
	 * @param Message - ���������
	 * @param Data - ���. ������
	 */
	this.Notify = function(Sender, Message, Data) {
		Log.Write(1, '[BTNotifyObject] Message: ' + Message + ", Data: " + Data);
	};
}

//-----------------------------------------------------------------------------
// DateTime - ������ � ������
//-----------------------------------------------------------------------------

/**
 * ���������� ��������� ���� � ��������
 * @return DateTime
 */
function GetServerDateTime() {
	return Connector.GetServerDateTime();
}

/**
 * ���������� ��������� ����
 */
function GetServerDate() {
	var DateTime = GetSystemDateTime();	
	var Result = ExtractDate(DateTime);
	return Result;
}

/**
 * ���������� ������ ���� �������: ����.�����.��� ����:������:�������
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
 * @return Arr {Array} [0] - ������ ���, [1] ����� ���
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
 * ����������  �������� ������.
 * @param MonthIndex - ����� ������ (��������� � ����)
 * @param Multipli - ���������
 * @param Launge - ���� �� ������� ������ �������� ������ (RU, UA)
 *		�� ��������� ������������ �������.
 * @return Name {String} - �������� ������
 */ 
function GetMonthName(MonthIndex, Launge, Multipli) {
	var MonthRU1 = ["������", "�������", "����", "������", "���", "����", "����",
  		"������", "��������", "�������", "������", "�������"];
  	var MonthRU2 = ["������", "�������", "�����", "������", "���", "����", "����",
  		"�������", "��������", "�������", "������", "�������"];
	
	var MonthUA1 = ["ѳ����", "�����", "��������", "������", "�������", "�������", 
		"������", "�������", "��������", "�������", "��������", "��������"];
	var MonthUA2 = ["ѳ���", "�����", "�������", "�����", "������", "������", 
		"�����", "������", "�������", "������", "���������", "������"];

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
 * @test - ��������� �������� ������
 */
function testGetMonthName() {
	Assert.startTest("GetMonthName");

	// �������
	var Month = GetMonthName(0, 'RU');
	Assert.equals("������", Month);
	
	Month = GetMonthName(11, 'RU');
	Assert.equals("�������", Month);

	Month = GetMonthName(13, 'RU');
	Assert.equals("�������", Month);

	Month = GetMonthName(0, 'RU', true);
	Assert.equals("������", Month);
	
	// ����������
	Month = GetMonthName(0, 'UA');
	Assert.equals("ѳ����", Month);
	
	Month = GetMonthName(11, 'UA');
	Assert.equals("��������", Month);
	
	Month = GetMonthName(13, 'UA');
	Assert.equals("�����", Month);

	Month = GetMonthName(0, 'UA', true);
	Assert.equals("ѳ���", Month);
	
	Assert.finishTest();
}

//-----------------------------------------------------------------------------
// Workflow - ������ �� ������ � ������ ����������
//
// ������������ �������: scr_WorkflowUtils
//-----------------------------------------------------------------------------

/**
 * ��������� ��, ������ ����������:
 * Parameters = 
 * @param WorkflowUSI - USI ��������� ��
 * @param Parameters - ������, ���������� �������� ��� ���� �� ����������:
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
 * ���������� workflow engine. � ������ ������ �� ������, ������� �����
 * ��������� �����
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
 * ���������� �������� ��������� �� �������� ��
 * @param Item - ������� ��������
 * @param ParameterName - ��� ��������� � ���������
 * @return Value - ��������
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
		BTCatchException("������ ��������� ��������� ��������: " + ParameterName);
	}
	return Value;
}

/**
 * ������������� � �������� ��������
 * @param Item - ������� ��������
 * @param ParameterName - ��� ��������� � ���������
 * @param Value - ��������
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
		BTCatchException('������ ��������� �������� ��������� ��������: ' + 
			ParameterName + ', ��������: ' + Value);
	}
}

//-----------------------------------------------------------------------------
// Right - ����� �������
//-----------------------------------------------------------------------------

/**
 * �������� ����� ������� ��� ������
 * @param TableUSI - ������� ���� (������: tbl_ContactRight
 * @param RecordID - ������ � ������� ����������� �������
 * @AdminUnitID - ID ������ �� �������� ds_AdminUnit
 * @param CanRead - ������ �� ������
 * @param CanWrite - ������ �� ��������������
 * @param CanChange - ������ �� ��������� �������
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
 * �������� ������/������������ �� �������� ��� ���� �������
 * @param TableUSI - ������� ���� (������: tbl_ContactRight)
 * @AdminUnitID - ID ������ �� �������� ds_AdminUnit 
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
 * �������� ���� ���� �������� ��� ���� ������� �������
 * @param TableUSI - ������� ���� (������: tbl_ContactRight)
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
 * �������� ������/������������ �� �������� ��� ����� ������
 * @param RecordID - ������, � ������� ����������� �������
 * @param TableUSI - ������� ���� (������: tbl_ContactRight)
 * @AdminUnitID - ID ������ �� �������� ds_AdminUnit 
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
 * ���������� ���� ������� ��� ����� ������
 * ���� ����������� ������ ����������, ����� ����������� �����
 * @param TableUSI - ������� ���� (������: tbl_ContactRight)
 * @AdminUnitID - ID ������ �� �������� ds_AdminUnit
 * @param CanRead - ������ �� ������
 * @param CanWrite - ������ �� ��������������
 * @param CanChange - ������ �� ��������� �������
 * @param IsGroup - ��� ������(true) ��� �������(false)
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
 * ���������� ���� ������� ���� ������� �� ��������.
 * ���� ����������� ������ ����������, ����� ����������� �����
 * @param DatasetUSI - ������� �������� ����� ��������� ������
 * @param TableUSI - ������� ���� (������: tbl_ContactRight)
 * @AdminUnitID - ID ������ �� �������� ds_AdminUnit
 * @param CanRead - ������ �� ������
 * @param CanWrite - ������ �� ��������������
 * @param CanChange - ������ �� ��������� �������
 * @param IsGroup - ��� ������(true) ��� �������(false)
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
// Files - ������ � �������
//
// ������������ �������: scr_FileUtils
//-----------------------------------------------------------------------------

function Files() {

	var _InsertQueryUSI = '';
	var _SubjectFieldName = '';
	
	/**
	 * @param Dataset - ������� ������
	 * @param ObjectType - ��� ����� (������, ����, ���� � �����)
	 * @param OldFileID
	 * @param ShortFileName - �������� ����� ��� ����������� � ������
	 * @param LongFileName - ������ ���� � �����
	 * @param Revision - ��������� 1
	 * @param AddedFileIDsArray - ������ ������
	 * @param SubjectID - ID ������ � ������� ����������� ����
	 * @param IsLog - ������� ��������� ID �����
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
	 * @param FileID - ������ �� ���� �� ������� ������
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
		// ��� ������������ ����� - ����.
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
	 * ��������� ��� ����� �� ��� ����
	 * @param FilePatch - ���� � �����
	 */
	ExtractFileName : function(FilePatch) {
		var FilePatchArr = FilePatch.split('\\');
		var FileName = FilePatchArr[FilePatchArr.length - 1];
		return FileName;
	}
};

/**
 * �������� ����� � ������ "�����"
 * @param FilePatch - ���� � ����� �� ����������
 * @param DisplayName - ������������ ���
 * @param InsertQueryUSI - ������ ����������
 * @param SubjectID - ID ������ � ������� ����� ����������� ���������� �����
 * 					  �������� � �������� ��� �����������
 * @param SubjectFieldName - �������� ���� � ������� �������� SubjectID
 */
function LoadDBFile(FilePatch, InsertQueryUSI, SubjectID, SubjectFieldName) {
	var DisplayName = DBFiles.ExtractFileName(FilePatch);
	
	DBFiles.LoadDBFile(FilePatch, DisplayName, InsertQueryUSI, 
		SubjectID, SubjectFieldName);	
}

//-----------------------------------------------------------------------------
// Exceptions - ���������� ������
//-----------------------------------------------------------------------------

/**
 * Exception �����, ���������� � ����� ������ ��������� ������ � ��� ���������
 * @param Exception - Exception object
 * @param Message - ������������ ������ �� ������
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
		ErrMessage = "����������� ������: " + Ex.toString();
	}
	
	var MessageStr = (Message ? 
		FormatStr('��������� ������ � ������ %1: %2, ������������ ��������� - %3', 
			Caller, Message, ErrMessage) :
		FormatStr('��������� ������ � ������ %1: %2', 
			Caller, ErrMessage));
	Log.Write(2, MessageStr);
}

// ----------------------------------------------------------------------------
// Caching system - ����������� ��������
// ----------------------------------------------------------------------------

/**
 * ������� �����������. �������� � ���� ���������� �������
 * ������������ ������ Singleton ��� ������ ���������� �������
 * ������� ������ ���� �����: �������, ��������, ��������� �������
 */
var CacheSystem = new function() {
	var instance;
		
	// ��� ��� ��������
	var CacheItem = {};
	
	// ��� ��� ���������
	var CacheDataset = {};
	
	// ��� ��� ��������
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
 * ���������� �� ���� �������.
 * ���� �� ������� - ������� ����� ������� � ����������.
 * @param DatasetUSI - �������� ��������
 * @param Key - ���� ������������
 * @return �������
 */
function DatasetCache_Get(DatasetUSI, Key) {	
	return CS.GetDataset(DatasetUSI, Key);		
}

/**
 * ���������� ������ � �������� ��� � �������.
 * @param ScriptUSI - USI �������
 * @return ������
 */
function GetScript(ScriptUSI) {
	return CS.GetScript(ScriptUSI);		
}

/**
 * �������� � ���������� �������
 * @ItemUSI - ������ � �������
 * @Key - ���� ������������
 * @return ������
 */
function CacheItem(ItemUSI, Key) {
	return CS.GetItem(ItemUSI, Key);
}

//-----------------------------------------------------------------------------
// Thunderbird - �������� ����� ����� �������� Thunderbird
//
// ������ �������������:
//
// ������� ����� ������
// var TH = new Thunderbird();
// TH.CreateNewEmail();
//
// ������� ����� � ����������� ������
// var TH = new Thunderbird();
// TH.SendSingleEmail("support@btech.com", "��� ����", "<b>��� ����� ���������</b>");
//-----------------------------------------------------------------------------

function Thunderbird() { }

Thunderbird.prototype = {
	constructor: Thunderbird,
		
	/**
	 * ������� ������ ���������
	 */
	CreateNewEmail: function()
	{
		var WshShell = new ActiveXObject("WScript.Shell");
		WshShell.Run('thunderbird -compose');
	},
	
	/**
	 *	To - ����� ���� ���������� (MyEmail@mail.com)
	 *	Subject - ���� ��������� (Dear friend!)
	 *	Body - ����� ��������� (Dear friend, i'm fine....)
	 *	AttachmentList - ������, ������� �������� ������ �� ����� � ����������
	 */
	SendSingleEmail: function(To, Subject, Body, AttachmentList)
	{
		var WshShell = new ActiveXObject("WScript.Shell");
		try {
		WshShell.Run("thunderbird -compose \"to='" + To + "\',subject='" + 
			Subject + "',body='" + Body + "'");
		} catch(e) {
			Log.Write(2, "��������� ������ ��� ��������, ��������� ���������� �� ��������. " + e.message);
		} finally {
			WshShell = System.EmptyValue;
		}
	}
};

//-----------------------------------------------------------------------------
// UnitTest - ������������ ��������
//
// ������ �������������
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
 * ����� ��� ������ � Ajax
 * @author KARTOH
 * @author KBaluh {baluh.kostya@gmail.com}
 * @date 03.08.2012
 * @version 1.2
 */
var Ajax = new function() {
	
	/**
	 * ������� ������
	 */
	var instance; 
	
	var REQUEST_STATUS_READY = 4;
	var REQUEST_STATUS_DONE = 200;
 
	/***
	����������� - ������������ ������ Singleton
	*/
	function Ajax() {
		if (!instance)
			instance = this;
		else return instance;
	}

	/**
	 * ��������� �������� {Array}
	 */
	this.requests = [];
	
	/**
	 * ����� ������ ���������� �������.
	 * ���� ���������� ������� ����, ������� �����
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
	 * ����� �������� �������
	 */
	this.create_request = function() {
		var request;
	
		try {
			// ��������� ������� IE >= 6 ������
			request = new ActiveXObject("Msxml2.XMLHTTP");
		}
		catch (Ex) {
			try {
				// ��������� ������� IE <= 5 ������
				request = new ActiveXObject("Microsoft.XMLHTTP");
			}
			catch (ExIE_5) {
				BTCatchException(ExIE_5,
					"[Ajax] ���������� ������� XMLHttpRequest");
				request = false;
			}
		}
		return request;
	};
	
	/**
	 * ����� �������� �������
	 * @param {String} url - ����� ��������
	 * @param {String} send - ������������ ���������
	 * @param {Function} callback - ���������� � �������� ������� ����� ��������
	 * @param {Object} head - ���������
	 * @param {Boolean} asunc - ��������� / ���������� ���������� ��������
	 * 							�� ��������� ����������
	 */
	this.send_request = function(url, send, callback, head, async) {
		var request = this.get_request();
		if (!request) {
			return false;
		}
		
		var http = send ? 'POST' : 'GET';
		request.open(http, url, Boolean(async));
		
		// ������������ ����������
		for(var key in head) {
			request.setRequestHeader(key,head[key]);
		}
		
		// �������� �������
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
 * ����� ��� GET ������� �� �������� ��������
 * ��������� ������ � ���������� ����� ��������
 * @param Url
 * @return Result {String} - ����� ��������
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
 * @test - ��������� �������� ������� � "null" head'����
 */
function testSendRequestNullHead() {

	/**
	 * ��������� �����, ������� ����� �� �������� �������
	 */
	var testAjaxCallbackMethod = function(ResponseText) {
		Log.Write(1, '������� ����� �������: ' + ResponseText);
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
	 * ������ � ������� ������ �� ������, ��� �� ���������� ��
	 */
	var ThrowInstallError = function() {
		BTCatchException("�� ������� �������� OpenOffice. " +
			"��������� ������������ ��������� � ������������� ����������� ActiveXObject.", 
			"[OpenOffice.ThrowInstallError]");				
	};
	
	/**
	 * ������������� � ������� �������� - ��������
	 * @param Document
	 */
	var SetActiveDocument = function(Document) {
		_Documents.push(Document);
		_ActiveDocument = Document;
	};

	/**
	 * ���������� ������ �� ������ � ��
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
	 * ��������� ����
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
	 * ������� ����� ��������� ��������
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
	 * ���������� �������� ��������
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
 * ������ ���������� ��� ������ ��� ���������
 */
function MessageThread() {
	System.ProcessMessages();
}

/**
 * ������ ������� ��� ������� "��������"
 */
function CursorWait() {
	System.BeginProcessing();	
}

/**
 * ������ ������� ��� ������� "�� ���������"
 */
function CursorDefault() {
	System.EndProcessing();	
}

// ----------------------------------------------------------------------------
// 1C - ������� ��� ������ � 1C
// ----------------------------------------------------------------------------

/**
 * ��������� ������� ����������
 * @param DataflowID - ID ����������
 * @param ItemID - ID �������� � ����������
 */
function RunOneCSyncItem(DataflowID, ItemID) {
	var Script = GetScript('scr_Dataflow1CUtils');
	Script.ScriptControl.Run('ExportObject', null, DataflowID, ItemID);  
}

/**
 * ����� ��� ����������� � ���� 1�
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
// Mock objects - ������� � ����������� ���������� ��������
// ----------------------------------------------------------------------------

/**
 * �������
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
 * ����
 */
function MockWindow() {
	this.Attributes	= GetNewDictionary();
	this.Attributes.Add('DefaultValues', GetNewDictionary());	
}

/**
 * ��������� ��
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
 * ��������� ����� ��
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
 * �������� ����� ��
 */
function MockParameterMap() {
	this.Name = "";
	this.Value = null;	
}

/**
 * ������� ��
 */
function MockActionItem() {
	this.ParametersMap = new MockParametersMap();
	this.Parameters = new MockParameters();
	this.ParentItems = new MockParameters();
	this.ParentItems.ParentDiagram = new MockDiagram();	
}

/**
 * ��������� ��
 */
function MockDiagram() {
	this.Parameters = new MockParameters();
}

//-----------------------------------------------------------------------------
// Init framework
//-----------------------------------------------------------------------------

// �����������
var CS = new CacheSystem();

// �������� ���� ������
var Assert = new UnitTest();

// �����
var DBFiles = new Files();

//-----------------------------------------------------------------------------
// Private methods - ���� �� ���������� scr_DB, ���� ������� ���������� ���
// ������ ������� �������
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
