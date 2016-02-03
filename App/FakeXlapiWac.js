if (typeof(OfficeExtension.WacRuntime.IAsyncEnumerator) == "undefined")
{
	OfficeExtension.WacRuntime.IAsyncEnumerator = function()
	{
	};
	OfficeExtension.WacRuntime.IAsyncEnumerator.prototype = {
	    get_current: null,
	    moveNext: null
	};
	OfficeExtension.WacRuntime.IAsyncEnumerator.registerInterface("OfficeExtension.WacRuntime.IAsyncEnumerator");
}


Type.registerNamespace('ExcelApiWac');

ExcelApiWac.ChartType = function() {}
ExcelApiWac.ChartType.prototype = {
    none: 0, 
    pie: 1, 
    bar: 2, 
    line: 3, 
    _3DBar: 4
}
ExcelApiWac.ChartType.registerEnum('ExcelApiWac.ChartType', false);


ExcelApiWac.ErrorMethodType = function() {}
ExcelApiWac.ErrorMethodType.prototype = {
    none: 0, 
    accessDenied: 1, 
    stateChanged: 2, 
    bounds: 3, 
    abort: 4
}
ExcelApiWac.ErrorMethodType.registerEnum('ExcelApiWac.ErrorMethodType', false);


ExcelApiWac.RangeValueType = function() {}
ExcelApiWac.RangeValueType.prototype = {
    unknown: 0, 
    empty: 1, 
    string: 2, 
    integer: 3, 
    double: 4, 
    boolean: 5, 
    error: 6
}
ExcelApiWac.RangeValueType.registerEnum('ExcelApiWac.RangeValueType', false);


ExcelApiWac.Application = function ExcelApiWac_Application() {
}
ExcelApiWac.Application.prototype = {
    
    get_activeWorkbook: function ExcelApiWac_Application$get_activeWorkbook$in() {
        return new ExcelApiWac.Workbook();
    },
    
    get_hasBase: function ExcelApiWac_Application$get_hasBase$in() {
        return false;
    },
    
    get_testWorkbook: function ExcelApiWac_Application$get_testWorkbook$in() {
        return new ExcelApiWac.TestWorkbook();
    },
    
    _GetObjectByReferenceId: function ExcelApiWac_Application$_GetObjectByReferenceId$in(referenceId) {
        var rangeData = ExcelApiWac._rangeData.getFromCache(referenceId);
        if (rangeData) {
            var range = new ExcelApiWac.Range('A1');
            range._setReferenceId$i$0(referenceId, rangeData);
            return range;
        }
        throw Error.argument();
    },
    
    _GetObjectTypeNameByReferenceId: function ExcelApiWac_Application$_GetObjectTypeNameByReferenceId$in(referenceId) {
        return 'ExcelApi.Range';
    },
    
    _RemoveReference: function ExcelApiWac_Application$_RemoveReference$in(referenceId) {
        ExcelApiWac._rangeData.removeFromCache(referenceId);
    },
    
    getTypeName: function ExcelApiWac_Application$getTypeName$in() {
        return 'ExcelApi.Application';
    },
    
    invokePropertyGet: function ExcelApiWac_Application$invokePropertyGet$in(dispatchId) {
        switch (dispatchId) {
            case ExcelApiWac.Application._dispiD_ActiveWorkbook$p:
                return this.get_activeWorkbook();
            case ExcelApiWac.Application._dispiD_HasBase$p:
                return this.get_hasBase();
            case ExcelApiWac.Application._dispiD_TestWorkbook$p:
                return this.get_testWorkbook();
        }
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokePropertyGetAsync: function ExcelApiWac_Application$invokePropertyGetAsync$in(dispatchId) {
        var ret = this.invokePropertyGet(dispatchId);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(ret);
    },
    
    invokePropertySet: function ExcelApiWac_Application$invokePropertySet$in(dispatchId, args) {
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokePropertySetAsync: function ExcelApiWac_Application$invokePropertySetAsync$in(dispatchId, args) {
        this.invokePropertySet(dispatchId, args);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(null);
    },
    
    invokeMethod: function ExcelApiWac_Application$invokeMethod$in(dispatchId, args) {
        switch (dispatchId) {
            case ExcelApiWac.Application._dispiD__GetObjectByReferenceId$p:
                return this._GetObjectByReferenceIdInvoke$0(args);
            case ExcelApiWac.Application._dispiD__GetObjectTypeNameByReferenceId$p:
                return this._GetObjectTypeNameByReferenceIdInvoke$0(args);
            case ExcelApiWac.Application._dispiD__RemoveReference$p:
                return this._RemoveReferenceInvoke$0(args);
        }
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokeMethodAsync: function ExcelApiWac_Application$invokeMethodAsync$in(dispatchId, args) {
        var ret = this.invokeMethod(dispatchId, args);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(ret);
    },
    
    _GetObjectByReferenceIdInvoke$0: function ExcelApiWac_Application$_GetObjectByReferenceIdInvoke$0$in(args) {
        var varReferenceId = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'string');
        return this._GetObjectByReferenceId(varReferenceId);
    },
    
    _GetObjectTypeNameByReferenceIdInvoke$0: function ExcelApiWac_Application$_GetObjectTypeNameByReferenceIdInvoke$0$in(args) {
        var varReferenceId = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'string');
        return this._GetObjectTypeNameByReferenceId(varReferenceId);
    },
    
    _RemoveReferenceInvoke$0: function ExcelApiWac_Application$_RemoveReferenceInvoke$0$in(args) {
        var varReferenceId = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'string');
        this._RemoveReference(varReferenceId);
        return null;
    }
}


ExcelApiWac.Chart = function ExcelApiWac_Chart(id) {
    this._m_id$p$0 = id;
    this._m_title$p$0 = 'Init Chart Title';
}
ExcelApiWac.Chart.prototype = {
    _m_id$p$0: 0,
    _m_chartType$p$0: 0,
    _m_title$p$0: null,
    _m_nullableShowLabel$p$0: null,
    _m_nullableChartType$p$0: null,
    
    get_chartType: function ExcelApiWac_Chart$get_chartType$in() {
        var ret = this._m_chartType$p$0;
        if (!ret) {
            ret = (this._m_id$p$0 % 5);
        }
        return ret;
    },
    
    set_chartType: function ExcelApiWac_Chart$set_chartType$in(value) {
        this._m_chartType$p$0 = value;
        return value;
    },
    
    get_id: function ExcelApiWac_Chart$get_id$in() {
        return this._m_id$p$0;
    },
    
    get_imageData: function ExcelApiWac_Chart$get_imageData$in() {
        var str = '';
        for (var i = 1; i < 1547; i++) {
            str += String.fromCharCode(i);
        }
        return str;
    },
    
    get_name: function ExcelApiWac_Chart$get_name$in() {
        return 'Chart' + this._m_id$p$0;
    },
    
    get_nullableChartType: function ExcelApiWac_Chart$get_nullableChartType$in() {
        if (!(this._m_id$p$0 % 5)) {
            return null;
        }
        return ExcelApiWac.ChartType.pie;
    },
    
    set_nullableChartType: function ExcelApiWac_Chart$set_nullableChartType$in(value) {
        this._m_nullableChartType$p$0 = value;
        return value;
    },
    
    get_nullableShowLabel: function ExcelApiWac_Chart$get_nullableShowLabel$in() {
        if (!(this._m_id$p$0 % 5)) {
            return null;
        }
        return true;
    },
    
    set_nullableShowLabel: function ExcelApiWac_Chart$set_nullableShowLabel$in(value) {
        this._m_nullableShowLabel$p$0 = value;
        return value;
    },
    
    get_title: function ExcelApiWac_Chart$get_title$in() {
        return this._m_title$p$0;
    },
    
    set_title: function ExcelApiWac_Chart$set_title$in(value) {
        this._m_title$p$0 = value;
        return value;
    },
    
    deleteObject: function ExcelApiWac_Chart$deleteObject$in() {
        return;
    },
    
    getAsImage: function ExcelApiWac_Chart$getAsImage$in(large) {
        return null;
    },
    
    getTypeName: function ExcelApiWac_Chart$getTypeName$in() {
        return 'ExcelApi.Chart';
    },
    
    invokePropertyGet: function ExcelApiWac_Chart$invokePropertyGet$in(dispatchId) {
        switch (dispatchId) {
            case ExcelApiWac.Chart._dispiD_ChartType$p:
                return this.get_chartType();
            case ExcelApiWac.Chart._dispiD_Id$p:
                return this.get_id();
            case ExcelApiWac.Chart._dispiD_ImageData$p:
                return this.get_imageData();
            case ExcelApiWac.Chart._dispiD_Name$p:
                return this.get_name();
            case ExcelApiWac.Chart._dispiD_NullableChartType$p:
                return this.get_nullableChartType();
            case ExcelApiWac.Chart._dispiD_NullableShowLabel$p:
                return this.get_nullableShowLabel();
            case ExcelApiWac.Chart._dispiD_Title$p:
                return this.get_title();
        }
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokePropertyGetAsync: function ExcelApiWac_Chart$invokePropertyGetAsync$in(dispatchId) {
        var ret = this.invokePropertyGet(dispatchId);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(ret);
    },
    
    invokePropertySet: function ExcelApiWac_Chart$invokePropertySet$in(dispatchId, args) {
        switch (dispatchId) {
            case ExcelApiWac.Chart._dispiD_ChartType$p:
                this.set_chartType(OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'int'));
                return;
            case ExcelApiWac.Chart._dispiD_NullableChartType$p:
                this.set_nullableChartType(OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'int?'));
                return;
            case ExcelApiWac.Chart._dispiD_NullableShowLabel$p:
                this.set_nullableShowLabel(OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'boolean?'));
                return;
            case ExcelApiWac.Chart._dispiD_Title$p:
                this.set_title(OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'string'));
                return;
        }
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokePropertySetAsync: function ExcelApiWac_Chart$invokePropertySetAsync$in(dispatchId, args) {
        this.invokePropertySet(dispatchId, args);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(null);
    },
    
    invokeMethod: function ExcelApiWac_Chart$invokeMethod$in(dispatchId, args) {
        switch (dispatchId) {
            case ExcelApiWac.Chart._dispiD_Delete$p:
                return this._deleteInvoke$p$0(args);
            case ExcelApiWac.Chart._dispiD_GetAsImage$p:
                return this._getAsImageInvoke$p$0(args);
        }
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokeMethodAsync: function ExcelApiWac_Chart$invokeMethodAsync$in(dispatchId, args) {
        var ret = this.invokeMethod(dispatchId, args);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(ret);
    },
    
    _deleteInvoke$p$0: function ExcelApiWac_Chart$_deleteInvoke$p$0$in(args) {
        this.deleteObject();
        return null;
    },
    
    _getAsImageInvoke$p$0: function ExcelApiWac_Chart$_getAsImageInvoke$p$0$in(args) {
        var varLarge = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'boolean');
        return this.getAsImage(varLarge);
    }
}


ExcelApiWac.ChartCollection = function ExcelApiWac_ChartCollection() {
}
ExcelApiWac.ChartCollection.prototype = {
    
    get_count: function ExcelApiWac_ChartCollection$get_count$in() {
        return 10;
    },
    
    add: function ExcelApiWac_ChartCollection$add$in(name, chartType) {
        var ret = new ExcelApiWac.Chart(1005);
        ret.set_chartType(chartType);
        return ret;
    },
    
    getItem: function ExcelApiWac_ChartCollection$getItem$in(index) {
        if (typeof(index) === 'number') {
            return new ExcelApiWac.Chart(index);
        }
        else if (typeof(index) === 'string') {
            var id = this._getChartIdFromName$p$0(index);
            return new ExcelApiWac.Chart(id);
        }
        else {
            throw Error.argument();
        }
    },
    
    getItemAt: function ExcelApiWac_ChartCollection$getItemAt$in(ordinal) {
        return new ExcelApiWac.Chart(ordinal + 1000);
    },
    
    getTypeName: function ExcelApiWac_ChartCollection$getTypeName$in() {
        return 'ExcelApi.ChartCollection';
    },
    
    invokePropertyGet: function ExcelApiWac_ChartCollection$invokePropertyGet$in(dispatchId) {
        switch (dispatchId) {
            case ExcelApiWac.ChartCollection._dispiD_Count$p:
                return this.get_count();
        }
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokePropertyGetAsync: function ExcelApiWac_ChartCollection$invokePropertyGetAsync$in(dispatchId) {
        var ret = this.invokePropertyGet(dispatchId);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(ret);
    },
    
    invokePropertySet: function ExcelApiWac_ChartCollection$invokePropertySet$in(dispatchId, args) {
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokePropertySetAsync: function ExcelApiWac_ChartCollection$invokePropertySetAsync$in(dispatchId, args) {
        this.invokePropertySet(dispatchId, args);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(null);
    },
    
    invokeMethod: function ExcelApiWac_ChartCollection$invokeMethod$in(dispatchId, args) {
        switch (dispatchId) {
            case ExcelApiWac.ChartCollection._dispiD_Add$p:
                return this._addInvoke$p$0(args);
            case ExcelApiWac.ChartCollection._dispiD_Item$p:
                return this._getItemInvoke$p$0(args);
            case ExcelApiWac.ChartCollection._dispiD_GetItemAt$p:
                return this._getItemAtInvoke$p$0(args);
        }
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokeMethodAsync: function ExcelApiWac_ChartCollection$invokeMethodAsync$in(dispatchId, args) {
        var ret = this.invokeMethod(dispatchId, args);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(ret);
    },
    
    _addInvoke$p$0: function ExcelApiWac_ChartCollection$_addInvoke$p$0$in(args) {
        var varName = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'string');
        var varChartType = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 1, 'int');
        return this.add(varName, varChartType);
    },
    
    _getItemInvoke$p$0: function ExcelApiWac_ChartCollection$_getItemInvoke$p$0$in(args) {
        var varIndex = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'any');
        return this.getItem(varIndex);
    },
    
    _getItemAtInvoke$p$0: function ExcelApiWac_ChartCollection$_getItemAtInvoke$p$0$in(args) {
        var varOrdinal = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'int');
        return this.getItemAt(varOrdinal);
    },
    
    _getChartIdFromName$p$0: function ExcelApiWac_ChartCollection$_getChartIdFromName$p$0$in(bstr) {
        if (bstr.substr(0, 'Chart'.length) === 'Chart') {
            return parseInt(bstr.substr('Chart'.length));
        }
        throw Error.argument();
    }
}


ExcelApiWac.TypeRegister = function ExcelApiWac_TypeRegister() {
}
ExcelApiWac.TypeRegister.registerTypes = function ExcelApiWac_TypeRegister$registerTypes$st(typeReg) {
    var typeInfo;
    var methodInfo;
    var propInfo;
    typeInfo = typeReg.addEnumType('ExcelApi.ChartType');
    typeInfo.addEnumField('None', 0);
    typeInfo.addEnumField('Pie', 1);
    typeInfo.addEnumField('Bar', 2);
    typeInfo.addEnumField('Line', 3);
    typeInfo.addEnumField('3DBar', 4);
    typeInfo = typeReg.addEnumType('ExcelApi.RangeValueType');
    typeInfo.addEnumField('Unknown', 0);
    typeInfo.addEnumField('Empty', 1);
    typeInfo.addEnumField('String', 2);
    typeInfo.addEnumField('Integer', 3);
    typeInfo.addEnumField('Double', 4);
    typeInfo.addEnumField('Boolean', 5);
    typeInfo.addEnumField('Error', 6);
    typeInfo = typeReg.addType('ExcelApi.Application', '{669eb674-96a9-431b-b26a-2f0b82b2542a}', '');
    propInfo = typeInfo.addProperty('ActiveWorkbook', 1, 'ExcelApi.Workbook', 3, true);
    propInfo = typeInfo.addProperty('HasBase', 6, 'Boolean', 1, true);
    propInfo = typeInfo.addProperty('TestWorkbook', 2, 'ExcelApi.TestWorkbook', 3, true);
    methodInfo = typeInfo.addMethod('_GetObjectByReferenceId', 4, 'Object', 1, 1, '_GetObjectByReferenceId');
    methodInfo.addParameter('referenceId', 'String', 1);
    methodInfo = typeInfo.addMethod('_GetObjectTypeNameByReferenceId', 3, 'String', 1, 1, '_GetObjectTypeNameByReferenceId');
    methodInfo.addParameter('referenceId', 'String', 1);
    methodInfo = typeInfo.addMethod('_RemoveReference', 5, 'Void', 0, 1, '_RemoveReference');
    methodInfo.addParameter('referenceId', 'String', 1);
    typeInfo = typeReg.addType('ExcelApi.Chart', '{d0447bab-634b-46ef-b9e9-c05fd58456f3}', '');
    propInfo = typeInfo.addProperty('ChartType', 2, 'ExcelApi.ChartType', 1, false);
    propInfo = typeInfo.addProperty('Id', 5, 'Int32', 1, true);
    propInfo = typeInfo.addProperty('ImageData', 8, 'String', 1, true);
    propInfo = typeInfo.addProperty('Name', 1, 'String', 1, true);
    propInfo = typeInfo.addProperty('NullableChartType', 6, 'ExcelApi.ChartType?', 1, false);
    propInfo = typeInfo.addProperty('NullableShowLabel', 7, 'Boolean?', 1, false);
    propInfo = typeInfo.addProperty('Title', 3, 'String', 1, false);
    methodInfo = typeInfo.addMethod('Delete', 4, 'Void', 0, 0, 'Delete');
    methodInfo = typeInfo.addMethod('GetAsImage', 9, 'Stream', 1, 1, 'AsImage');
    methodInfo.addParameter('large', 'Boolean', 1);
    typeInfo.setRESTfulOperationMethodNames('', 'Delete');
    typeInfo = typeReg.addType('ExcelApi.ChartCollection', '{d22ddf8f-a9a0-4b2a-b860-d7db0ff9eccc}', '');
    propInfo = typeInfo.addProperty('Count', 1, 'Int32', 1, true);
    methodInfo = typeInfo.addMethod('Add', 3, 'ExcelApi.Chart', 3, 0, 'Add');
    methodInfo.addParameter('name', 'String', 1);
    methodInfo.addParameter('chartType', 'ExcelApi.ChartType', 1);
    methodInfo = typeInfo.addIndexerMethod(2, 'ExcelApi.Chart', 3);
    methodInfo.addParameter('index', 'Object', 1);
    methodInfo = typeInfo.addMethod('GetItemAt', 4, 'ExcelApi.Chart', 3, 1, 'ItemAt');
    methodInfo.addParameter('ordinal', 'Int32', 1);
    typeInfo.setRESTfulOperationMethodNames('Add', '');
    typeInfo.setCollectionInfo('ExcelApi.Chart', false);
    typeInfo = typeReg.addEnumType('ExcelApi.ErrorMethodType');
    typeInfo.addEnumField('None', 0);
    typeInfo.addEnumField('AccessDenied', 1);
    typeInfo.addEnumField('StateChanged', 2);
    typeInfo.addEnumField('Bounds', 3);
    typeInfo.addEnumField('Abort', 4);
    typeInfo = typeReg.addType('ExcelApi.TestWorkbook', '{9d460404-bcf7-4f46-99d5-f07e7c4ad2de}', '');
    propInfo = typeInfo.addProperty('ErrorWorksheet', 1, 'ExcelApi.Worksheet', 3, true);
    propInfo = typeInfo.addProperty('ErrorWorksheet2', 5, 'ExcelApi.Worksheet', 3, true);
    methodInfo = typeInfo.addMethod('ErrorMethod', 2, 'String', 1, 0, 'ErrorMethod');
    methodInfo.addParameter('input', 'ExcelApi.ErrorMethodType', 1);
    methodInfo = typeInfo.addMethod('ErrorMethod2', 6, 'String', 1, 0, 'ErrorMethod2');
    methodInfo.addParameter('input', 'ExcelApi.ErrorMethodType', 1);
    methodInfo = typeInfo.addMethod('GetActiveWorksheet', 4, 'ExcelApi.Worksheet', 3, 1, 'ActiveWorksheet');
    methodInfo = typeInfo.addMethod('GetCachedObjectCount', 10, 'Int32', 1, 0, 'CachedObjectCount');
    methodInfo = typeInfo.addMethod('GetNullableBoolValue', 9, 'Boolean?', 1, 0, 'NullableBoolValue');
    methodInfo.addParameter('nullable', 'Boolean', 1);
    methodInfo = typeInfo.addMethod('GetNullableEnumValue', 8, 'ExcelApi.ChartType?', 1, 0, 'NullableEnumValue');
    methodInfo.addParameter('nullable', 'Boolean', 1);
    methodInfo = typeInfo.addMethod('GetObjectCount', 3, 'Int32', 1, 0, 'ObjectCount');
    methodInfo = typeInfo.addMethod('TestNullableInputValue', 7, 'String', 1, 0, 'TestNullableInputValue');
    methodInfo.addParameter('chartType', 'ExcelApi.ChartType?', 1);
    methodInfo.addParameter('boolValue', 'Boolean?', 1);
    typeInfo = typeReg.addType('ExcelApi.Workbook', '{bb02266c-6204-4e0d-baa3-cc1a928f573e}', '');
    propInfo = typeInfo.addProperty('ActiveWorksheet', 2, 'ExcelApi.Worksheet', 3, true);
    propInfo = typeInfo.addProperty('Charts', 3, 'ExcelApi.ChartCollection', 3, true);
    propInfo = typeInfo.addProperty('Sheets', 1, 'ExcelApi.WorksheetCollection', 3, true);
    methodInfo = typeInfo.addMethod('GetChartByType', 4, 'ExcelApi.Chart', 3, 1, 'ChartByType');
    methodInfo.addParameter('chartType', 'ExcelApi.ChartType', 1);
    methodInfo = typeInfo.addMethod('GetChartByTypeTitle', 5, 'ExcelApi.Chart', 3, 1, 'ChartByTypeTitle');
    methodInfo.addParameter('chartType', 'ExcelApi.ChartType', 1);
    methodInfo.addParameter('title', 'String', 1);
    methodInfo = typeInfo.addMethod('SomeAction', 6, 'String', 1, 0, 'SomeAction');
    methodInfo.addParameter('intVal', 'Int32', 1);
    methodInfo.addParameter('strVal', 'String', 1);
    methodInfo.addParameter('enumVal', 'ExcelApi.ChartType', 1);
    typeInfo = typeReg.addType('ExcelApi.Worksheet', '{b86e5ae1-476e-4e56-825d-885468e549f3}', '');
    propInfo = typeInfo.addProperty('ActiveCell', 4, 'ExcelApi.Range', 3, true);
    propInfo = typeInfo.addProperty('ActiveCellInvalidAfterRequest', 5, 'ExcelApi.Range', 3, true);
    propInfo = typeInfo.addProperty('CalculatedName', 7, 'String', 1, false);
    propInfo = typeInfo.addProperty('Name', 3, 'String', 1, true);
    propInfo = typeInfo.addProperty('_Id', 6, 'Int32', 1, true);
    methodInfo = typeInfo.addMethod('NullChart', 9, 'ExcelApi.Chart', 3, 1, 'NullChart');
    methodInfo.addParameter('address', 'String', 1);
    methodInfo = typeInfo.addMethod('NullRange', 8, 'ExcelApi.Range', 3, 1, 'NullRange');
    methodInfo.addParameter('address', 'String', 1);
    methodInfo = typeInfo.addMethod('Range', 1, 'ExcelApi.Range', 3, 1, 'Range');
    methodInfo.addParameter('address', 'String', 1);
    methodInfo = typeInfo.addMethod('SomeRangeOperation', 2, 'String', 1, 0, 'SomeRangeOperation');
    methodInfo.addParameter('input', 'String', 1);
    methodInfo.addParameter('range', 'ExcelApi.Range', 3);
    methodInfo = typeInfo.addMethod('_RestOnly', 10, 'String', 1, 1, 'RestOnlyOperation');
    typeInfo = typeReg.addType('ExcelApi.WorksheetCollection', '{55a36c77-3310-4afb-aa64-3c1a685f2f50}', '');
    methodInfo = typeInfo.addMethod('Add', 5, 'ExcelApi.Worksheet', 3, 0, 'Add');
    methodInfo.addParameter('name', 'String', 1);
    methodInfo = typeInfo.addMethod('FindSheet', 4, 'ExcelApi.Worksheet', 3, 1, 'FindSheet');
    methodInfo.addParameter('text', 'String', 1);
    methodInfo = typeInfo.addMethod('GetActiveWorksheetInvalidAfterRequest', 2, 'ExcelApi.Worksheet', 3, 1, 'ActiveWorksheetInvalidAfterRequest');
    methodInfo = typeInfo.addMethod('GetItem', 3, 'ExcelApi.Worksheet', 3, 0, 'Item');
    methodInfo.addParameter('index', 'Object', 1);
    methodInfo = typeInfo.addIndexerMethod(1, 'ExcelApi.Worksheet', 3);
    methodInfo.addParameter('index', 'Object', 1);
    typeInfo.setRESTfulOperationMethodNames('Add', '');
    typeInfo.setCollectionInfo('ExcelApi.Worksheet', true);
    typeInfo = typeReg.addType('ExcelApi.Range', '{906962e8-a18a-4cc9-9342-279f056bc293}', '');
    propInfo = typeInfo.addProperty('ColumnIndex', 2, 'Int32', 1, true);
    propInfo = typeInfo.addProperty('LogText', 9, 'String', 1, true);
    propInfo = typeInfo.addProperty('RowIndex', 1, 'Int32', 1, true);
    propInfo = typeInfo.addProperty('Sort', 18, 'ExcelApi.RangeSort', 3, true);
    propInfo = typeInfo.addProperty('Text', 6, 'String', 1, false);
    propInfo = typeInfo.addProperty('TextArray', 8, 'String[][]', 1, false);
    propInfo = typeInfo.addProperty('Value', 3, 'Object', 1, false);
    propInfo = typeInfo.addProperty('ValueArray', 7, 'Object[][]', 1, false);
    propInfo = typeInfo.addProperty('ValueArray2', 12, 'Object', 1, false);
    propInfo = typeInfo.addProperty('ValueTypes', 16, 'ExcelApi.RangeValueType[][]', 1, true);
    propInfo = typeInfo.addProperty('_ReferenceId', 10, 'String', 1, true);
    methodInfo = typeInfo.addMethod('Activate', 5, 'Void', 0, 0, 'Activate');
    methodInfo = typeInfo.addMethod('GetValueArray2', 13, 'Object', 1, 0, 'ValueArray2');
    methodInfo = typeInfo.addMethod('NotRestMethod', 17, 'Int32', 1, 1, '');
    methodInfo = typeInfo.addMethod('ReplaceValue', 4, 'Object', 1, 0, 'ReplaceValue');
    methodInfo.addParameter('newValue', 'Object', 1);
    methodInfo = typeInfo.addMethod('SetValueArray2', 14, 'Void', 0, 0, 'SetValueArray2');
    methodInfo.addParameter('valueArray', 'Object', 1);
    methodInfo.addParameter('text', 'String', 1);
    methodInfo = typeInfo.addMethod('_KeepReference', 11, 'Void', 0, 1, '_KeepReference');
    methodInfo = typeInfo.addMethod('_OnAccess', 15, 'Void', 0, 1, '_OnAccess');
    typeInfo = typeReg.addType('ExcelApi.TestCaseObject', '{53984705-84a6-4393-87f3-a118cc7bb047}', '{4c4fd77f-8fb3-4e12-a796-84fde62d4eda}');
    methodInfo = typeInfo.addMethod('CalculateAddressAndSaveToRange', 1, 'String', 1, 0, 'CalculateAddressAndSaveToRange');
    methodInfo.addParameter('street', 'String', 1);
    methodInfo.addParameter('city', 'String', 1);
    methodInfo.addParameter('range', 'ExcelApi.Range', 3);
    methodInfo = typeInfo.addMethod('MatrixSum', 11, 'Double', 1, 0, 'MatrixSum');
    methodInfo.addParameter('matrix', 'Object[][]', 1);
    methodInfo = typeInfo.addMethod('Sum', 10, 'Double', 1, 0, 'Sum');
    methodInfo.addParameter('values', 'Object[]', 1);
    methodInfo = typeInfo.addMethod('TestParamBool', 2, 'Boolean', 1, 0, 'TestParamBool');
    methodInfo.addParameter('value', 'Boolean', 1);
    methodInfo = typeInfo.addMethod('TestParamDouble', 5, 'Double', 1, 0, 'TestParamDouble');
    methodInfo.addParameter('value', 'Double', 1);
    methodInfo = typeInfo.addMethod('TestParamFloat', 4, 'Single', 1, 0, 'TestParamFloat');
    methodInfo.addParameter('value', 'Single', 1);
    methodInfo = typeInfo.addMethod('TestParamInt', 3, 'Int32', 1, 0, 'TestParamInt');
    methodInfo.addParameter('value', 'Int32', 1);
    methodInfo = typeInfo.addMethod('TestParamRange', 7, 'ExcelApi.Range', 3, 0, 'TestParamRange');
    methodInfo.addParameter('value', 'ExcelApi.Range', 3);
    methodInfo = typeInfo.addMethod('TestParamString', 6, 'String', 1, 0, 'TestParamString');
    methodInfo.addParameter('value', 'String', 1);
    methodInfo = typeInfo.addMethod('TestUrlKeyValueDecode', 8, 'String', 1, 0, 'TestUrlKeyValueDecode');
    methodInfo.addParameter('value', 'String', 1);
    methodInfo = typeInfo.addMethod('TestUrlPathEncode', 9, 'String', 1, 0, 'TestUrlPathEncode');
    methodInfo.addParameter('value', 'String', 1);
    typeInfo = typeReg.addType('ExcelApi.RangeSort', '{b0eb6d54-c934-479c-974c-bf872c924656}', '');
    propInfo = typeInfo.addProperty('Fields', 2, 'ExcelApi.SortField[]', 2, true);
    propInfo = typeInfo.addProperty('Fields2', 7, 'Object[]', 1, false);
    propInfo = typeInfo.addProperty('QueryField', 4, 'ExcelApi.QueryWithSortField', 2, true);
    methodInfo = typeInfo.addMethod('Apply', 1, 'String', 1, 0, 'Apply');
    methodInfo.addParameter('fields', 'ExcelApi.SortField[]', 2);
    methodInfo = typeInfo.addMethod('ApplyAndReturnFirstField', 3, 'ExcelApi.SortField', 2, 0, 'ApplyAndReturnFirstField');
    methodInfo.addParameter('fields', 'ExcelApi.SortField[]', 2);
    methodInfo = typeInfo.addMethod('ApplyMixed', 6, 'Int32', 1, 0, 'ApplyMixed');
    methodInfo.addParameter('fields', 'Object[]', 1);
    methodInfo = typeInfo.addMethod('ApplyQueryWithSortFieldAndReturnLast', 5, 'ExcelApi.QueryWithSortField', 2, 0, 'ApplyQueryWithSortFieldAndReturnLast');
    methodInfo.addParameter('fields', 'ExcelApi.QueryWithSortField[]', 2);
    typeInfo = typeReg.addType('ExcelApi.SortField', '{337ff52d-6bd6-47b6-8047-d1a57f5abed4}', '{8c274050-f0bd-4ac1-8a1f-3b67ffd4c29d}');
    propInfo = typeInfo.addProperty('Assending', 2, 'Boolean', 1, false);
    propInfo = typeInfo.addProperty('ColumnIndex', 1, 'Int32', 1, false);
    propInfo = typeInfo.addProperty('UseCurrentCulture', 3, 'Boolean', 1, false);
    typeInfo = typeReg.addType('ExcelApi.QueryWithSortField', '{a4576ceb-16eb-42aa-9c96-ff9221db7699}', '{5267b183-56a2-4cc3-a0a1-5511e9564024}');
    propInfo = typeInfo.addProperty('Field', 2, 'ExcelApi.SortField', 2, false);
    propInfo = typeInfo.addProperty('RowLimit', 1, 'Int32', 1, false);
}


ExcelApiWac.QueryWithSortField = function ExcelApiWac_QueryWithSortField() {
}
ExcelApiWac.QueryWithSortField.prototype = {
    
    get_field: function ExcelApiWac_QueryWithSortField$get_field$in() {
        return null;
    },
    
    set_field: function ExcelApiWac_QueryWithSortField$set_field$in(value) {
        return value;
    },
    
    get_rowLimit: function ExcelApiWac_QueryWithSortField$get_rowLimit$in() {
        return 0;
    },
    
    set_rowLimit: function ExcelApiWac_QueryWithSortField$set_rowLimit$in(value) {
        return value;
    },
    
    getTypeName: function ExcelApiWac_QueryWithSortField$getTypeName$in() {
        return 'ExcelApi.QueryWithSortField';
    },
    
    invokePropertyGet: function ExcelApiWac_QueryWithSortField$invokePropertyGet$in(dispatchId) {
        switch (dispatchId) {
            case ExcelApiWac.QueryWithSortField._dispiD_Field$p:
                return this.get_field();
            case ExcelApiWac.QueryWithSortField._dispiD_RowLimit$p:
                return this.get_rowLimit();
        }
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokePropertyGetAsync: function ExcelApiWac_QueryWithSortField$invokePropertyGetAsync$in(dispatchId) {
        var ret = this.invokePropertyGet(dispatchId);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(ret);
    },
    
    invokePropertySet: function ExcelApiWac_QueryWithSortField$invokePropertySet$in(dispatchId, args) {
        switch (dispatchId) {
            case ExcelApiWac.QueryWithSortField._dispiD_Field$p:
                this.set_field(OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'ExcelApiWac.SortField'));
                return;
            case ExcelApiWac.QueryWithSortField._dispiD_RowLimit$p:
                this.set_rowLimit(OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'int'));
                return;
        }
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokePropertySetAsync: function ExcelApiWac_QueryWithSortField$invokePropertySetAsync$in(dispatchId, args) {
        this.invokePropertySet(dispatchId, args);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(null);
    },
    
    invokeMethod: function ExcelApiWac_QueryWithSortField$invokeMethod$in(dispatchId, args) {
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokeMethodAsync: function ExcelApiWac_QueryWithSortField$invokeMethodAsync$in(dispatchId, args) {
        var ret = this.invokeMethod(dispatchId, args);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(ret);
    }
}


ExcelApiWac.Range = function ExcelApiWac_Range(address) {
    this._m_rangeData$p$0 = new ExcelApiWac._rangeData();
    this._m_rangeData$p$0.set_value('Initial Value');
    this._m_address$p$0 = address;
}
ExcelApiWac.Range.prototype = {
    _m_address$p$0: null,
    _m_rangeData$p$0: null,
    _m_logText$p$0: null,
    _m_referenceId$p$0: null,
    
    _setReferenceId$i$0: function ExcelApiWac_Range$_setReferenceId$i$0$in(referenceId, rangeData) {
        this._m_referenceId$p$0 = referenceId;
        this._m_rangeData$p$0 = rangeData;
    },
    
    get_columnIndex: function ExcelApiWac_Range$get_columnIndex$in() {
        return 0;
    },
    
    get_logText: function ExcelApiWac_Range$get_logText$in() {
        return this._m_logText$p$0;
    },
    
    get_rowIndex: function ExcelApiWac_Range$get_rowIndex$in() {
        return 0;
    },
    
    get_sort: function ExcelApiWac_Range$get_sort$in() {
        return null;
    },
    
    get_text: function ExcelApiWac_Range$get_text$in() {
        return this._m_rangeData$p$0.get_text();
    },
    
    set_text: function ExcelApiWac_Range$set_text$in(value) {
        this._m_logText$p$0 += 'put_Text';
        this._m_rangeData$p$0.set_text(value);
        return value;
    },
    
    get_textArray: function ExcelApiWac_Range$get_textArray$in() {
        return this._m_rangeData$p$0.get_textArray();
    },
    
    set_textArray: function ExcelApiWac_Range$set_textArray$in(value) {
        this._m_logText$p$0 += 'put_TextArray';
        this._m_rangeData$p$0.set_textArray(value);
        return value;
    },
    
    get_value: function ExcelApiWac_Range$get_value$in() {
        return this._m_rangeData$p$0.get_value();
    },
    
    set_value: function ExcelApiWac_Range$set_value$in(value) {
        this._m_logText$p$0 += 'put_Value';
        this._m_rangeData$p$0.set_value(value);
        return value;
    },
    
    get_valueArray: function ExcelApiWac_Range$get_valueArray$in() {
        return this._m_rangeData$p$0.get_valueArray();
    },
    
    set_valueArray: function ExcelApiWac_Range$set_valueArray$in(value) {
        this._m_logText$p$0 += 'put_ValueArray';
        this._m_rangeData$p$0.set_valueArray(value);
        return value;
    },
    
    get_valueArray2: function ExcelApiWac_Range$get_valueArray2$in() {
        return this.get_valueArray();
    },
    
    set_valueArray2: function ExcelApiWac_Range$set_valueArray2$in(value) {
        this.set_valueArray(value);
        return value;
    },
    
    get_valueTypes: function ExcelApiWac_Range$get_valueTypes$in() {
        var size = 3;
        var ret = [];
        for (var i = 0; i < size; i++) {
            var item = [];
            for (var j = 0; j < size; j++) {
                item.push(j);
            }
            ret.push(item);
        }
        return (ret);
    },
    
    get__ReferenceId: function ExcelApiWac_Range$get__ReferenceId$in() {
        return this._m_referenceId$p$0;
    },
    
    activate: function ExcelApiWac_Range$activate$in() {
        return;
    },
    
    getValueArray2: function ExcelApiWac_Range$getValueArray2$in() {
        return this.get_valueArray();
    },
    
    notRestMethod: function ExcelApiWac_Range$notRestMethod$in() {
        return 1234;
    },
    
    replaceValue: function ExcelApiWac_Range$replaceValue$in(newValue) {
        var ret = this._m_rangeData$p$0.get_value();
        this._m_rangeData$p$0.set_value(newValue);
        return ret;
    },
    
    setValueArray2: function ExcelApiWac_Range$setValueArray2$in(valueArray, text) {
        this.set_valueArray(valueArray);
        this.set_text(text);
    },
    
    _KeepReference: function ExcelApiWac_Range$_KeepReference$in() {
        if (!this._m_referenceId$p$0) {
            this._m_referenceId$p$0 = ExcelApiWac._rangeData.putIntoCache(this._m_rangeData$p$0);
        }
    },
    
    _OnAccess: function ExcelApiWac_Range$_OnAccess$in() {
        this._m_logText$p$0 += '_OnAccess';
    },
    
    getTypeName: function ExcelApiWac_Range$getTypeName$in() {
        return 'ExcelApi.Range';
    },
    
    invokePropertyGet: function ExcelApiWac_Range$invokePropertyGet$in(dispatchId) {
        switch (dispatchId) {
            case ExcelApiWac.Range._dispiD_ColumnIndex$p:
                return this.get_columnIndex();
            case ExcelApiWac.Range._dispiD_LogText$p:
                return this.get_logText();
            case ExcelApiWac.Range._dispiD_RowIndex$p:
                return this.get_rowIndex();
            case ExcelApiWac.Range._dispiD_Sort$p:
                return this.get_sort();
            case ExcelApiWac.Range._dispiD_Text$p:
                return this.get_text();
            case ExcelApiWac.Range._dispiD_TextArray$p:
                return this.get_textArray();
            case ExcelApiWac.Range._dispiD_Value$p:
                return this.get_value();
            case ExcelApiWac.Range._dispiD_ValueArray$p:
                return this.get_valueArray();
            case ExcelApiWac.Range._dispiD_ValueArray2$p:
                return this.get_valueArray2();
            case ExcelApiWac.Range._dispiD_ValueTypes$p:
                return this.get_valueTypes();
            case ExcelApiWac.Range._dispiD__ReferenceId$p:
                return this.get__ReferenceId();
        }
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokePropertyGetAsync: function ExcelApiWac_Range$invokePropertyGetAsync$in(dispatchId) {
        var ret = this.invokePropertyGet(dispatchId);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(ret);
    },
    
    invokePropertySet: function ExcelApiWac_Range$invokePropertySet$in(dispatchId, args) {
        switch (dispatchId) {
            case ExcelApiWac.Range._dispiD_Text$p:
                this.set_text(OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'string'));
                return;
            case ExcelApiWac.Range._dispiD_TextArray$p:
                this.set_textArray(OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'Array<Array<string>>'));
                return;
            case ExcelApiWac.Range._dispiD_Value$p:
                this.set_value(OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'any'));
                return;
            case ExcelApiWac.Range._dispiD_ValueArray$p:
                this.set_valueArray(OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'Array<Array<any>>'));
                return;
            case ExcelApiWac.Range._dispiD_ValueArray2$p:
                this.set_valueArray2(OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'any'));
                return;
        }
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokePropertySetAsync: function ExcelApiWac_Range$invokePropertySetAsync$in(dispatchId, args) {
        this.invokePropertySet(dispatchId, args);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(null);
    },
    
    invokeMethod: function ExcelApiWac_Range$invokeMethod$in(dispatchId, args) {
        switch (dispatchId) {
            case ExcelApiWac.Range._dispiD_Activate$p:
                return this._activateInvoke$p$0(args);
            case ExcelApiWac.Range._dispiD_GetValueArray2$p:
                return this._getValueArray2Invoke$p$0(args);
            case ExcelApiWac.Range._dispiD_NotRestMethod$p:
                return this._notRestMethodInvoke$p$0(args);
            case ExcelApiWac.Range._dispiD_ReplaceValue$p:
                return this._replaceValueInvoke$p$0(args);
            case ExcelApiWac.Range._dispiD_SetValueArray2$p:
                return this._setValueArray2Invoke$p$0(args);
            case ExcelApiWac.Range._dispiD__KeepReference$p:
                return this._KeepReferenceInvoke$0(args);
            case ExcelApiWac.Range._dispiD__OnAccess$p:
                return this._OnAccessInvoke$0(args);
        }
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokeMethodAsync: function ExcelApiWac_Range$invokeMethodAsync$in(dispatchId, args) {
        var ret = this.invokeMethod(dispatchId, args);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(ret);
    },
    
    _activateInvoke$p$0: function ExcelApiWac_Range$_activateInvoke$p$0$in(args) {
        this.activate();
        return null;
    },
    
    _getValueArray2Invoke$p$0: function ExcelApiWac_Range$_getValueArray2Invoke$p$0$in(args) {
        return this.getValueArray2();
    },
    
    _notRestMethodInvoke$p$0: function ExcelApiWac_Range$_notRestMethodInvoke$p$0$in(args) {
        return this.notRestMethod();
    },
    
    _replaceValueInvoke$p$0: function ExcelApiWac_Range$_replaceValueInvoke$p$0$in(args) {
        var varNewValue = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'any');
        return this.replaceValue(varNewValue);
    },
    
    _setValueArray2Invoke$p$0: function ExcelApiWac_Range$_setValueArray2Invoke$p$0$in(args) {
        var varValueArray = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'any');
        var varText = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 1, 'string');
        this.setValueArray2(varValueArray, varText);
        return null;
    },
    
    _KeepReferenceInvoke$0: function ExcelApiWac_Range$_KeepReferenceInvoke$0$in(args) {
        this._KeepReference();
        return null;
    },
    
    _OnAccessInvoke$0: function ExcelApiWac_Range$_OnAccessInvoke$0$in(args) {
        this._OnAccess();
        return null;
    }
}


ExcelApiWac._rangeData = function ExcelApiWac__rangeData() {
}
ExcelApiWac._rangeData.putIntoCache = function ExcelApiWac__rangeData$putIntoCache$st(data) {
    if (!ExcelApiWac._rangeData._s_cache$p) {
        ExcelApiWac._rangeData._s_cache$p = {};
    }
    var key = 'R' + ExcelApiWac._rangeData._s_index$p;
    ExcelApiWac._rangeData._s_index$p++;
    ExcelApiWac._rangeData._s_cache$p[key] = data;
    return key;
}
ExcelApiWac._rangeData.getFromCache = function ExcelApiWac__rangeData$getFromCache$st(referenceId) {
    if (!ExcelApiWac._rangeData._s_cache$p) {
        ExcelApiWac._rangeData._s_cache$p = {};
    }
    return ExcelApiWac._rangeData._s_cache$p[referenceId];
}
ExcelApiWac._rangeData.removeFromCache = function ExcelApiWac__rangeData$removeFromCache$st(referenceId) {
    if (!ExcelApiWac._rangeData._s_cache$p) {
        ExcelApiWac._rangeData._s_cache$p = {};
    }
    delete ExcelApiWac._rangeData._s_cache$p[referenceId];
}
ExcelApiWac._rangeData.getCachedObjectCount = function ExcelApiWac__rangeData$getCachedObjectCount$st() {
    if (!ExcelApiWac._rangeData._s_cache$p) {
        ExcelApiWac._rangeData._s_cache$p = {};
    }
    var count = 0;
    var $$dict_2 = ExcelApiWac._rangeData._s_cache$p;
    for (var $$key_3 in $$dict_2) {
        var entry = { key: $$key_3, value: $$dict_2[$$key_3] };
        count++;
    }
    return count;
}
ExcelApiWac._rangeData.prototype = {
    _m_value$p$0: null,
    _m_text$p$0: null,
    _m_valueArray$p$0: null,
    _m_textArray$p$0: null,
    
    get_value: function ExcelApiWac__rangeData$get_value$in() {
        return this._m_value$p$0;
    },
    
    set_value: function ExcelApiWac__rangeData$set_value$in(value) {
        this._m_value$p$0 = value;
        return value;
    },
    
    get_text: function ExcelApiWac__rangeData$get_text$in() {
        return this._m_text$p$0;
    },
    
    set_text: function ExcelApiWac__rangeData$set_text$in(value) {
        this._m_text$p$0 = value;
        return value;
    },
    
    get_valueArray: function ExcelApiWac__rangeData$get_valueArray$in() {
        return this._m_valueArray$p$0;
    },
    
    set_valueArray: function ExcelApiWac__rangeData$set_valueArray$in(value) {
        this._m_valueArray$p$0 = value;
        return value;
    },
    
    get_textArray: function ExcelApiWac__rangeData$get_textArray$in() {
        return this._m_textArray$p$0;
    },
    
    set_textArray: function ExcelApiWac__rangeData$set_textArray$in(value) {
        this._m_textArray$p$0 = value;
        return value;
    }
}


ExcelApiWac.RangeSort = function ExcelApiWac_RangeSort() {
}
ExcelApiWac.RangeSort.prototype = {
    
    get_fields: function ExcelApiWac_RangeSort$get_fields$in() {
        return null;
    },
    
    get_fields2: function ExcelApiWac_RangeSort$get_fields2$in() {
        return null;
    },
    
    set_fields2: function ExcelApiWac_RangeSort$set_fields2$in(value) {
        return value;
    },
    
    get_queryField: function ExcelApiWac_RangeSort$get_queryField$in() {
        return null;
    },
    
    apply: function ExcelApiWac_RangeSort$apply$in(fields) {
        return null;
    },
    
    applyAndReturnFirstField: function ExcelApiWac_RangeSort$applyAndReturnFirstField$in(fields) {
        return null;
    },
    
    applyMixed: function ExcelApiWac_RangeSort$applyMixed$in(fields) {
        return 0;
    },
    
    applyQueryWithSortFieldAndReturnLast: function ExcelApiWac_RangeSort$applyQueryWithSortFieldAndReturnLast$in(fields) {
        return null;
    },
    
    getTypeName: function ExcelApiWac_RangeSort$getTypeName$in() {
        return 'ExcelApi.RangeSort';
    },
    
    invokePropertyGet: function ExcelApiWac_RangeSort$invokePropertyGet$in(dispatchId) {
        switch (dispatchId) {
            case ExcelApiWac.RangeSort._dispiD_Fields$p:
                return this.get_fields();
            case ExcelApiWac.RangeSort._dispiD_Fields2$p:
                return this.get_fields2();
            case ExcelApiWac.RangeSort._dispiD_QueryField$p:
                return this.get_queryField();
        }
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokePropertyGetAsync: function ExcelApiWac_RangeSort$invokePropertyGetAsync$in(dispatchId) {
        var ret = this.invokePropertyGet(dispatchId);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(ret);
    },
    
    invokePropertySet: function ExcelApiWac_RangeSort$invokePropertySet$in(dispatchId, args) {
        switch (dispatchId) {
            case ExcelApiWac.RangeSort._dispiD_Fields2$p:
                this.set_fields2(OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'Array<any>'));
                return;
        }
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokePropertySetAsync: function ExcelApiWac_RangeSort$invokePropertySetAsync$in(dispatchId, args) {
        this.invokePropertySet(dispatchId, args);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(null);
    },
    
    invokeMethod: function ExcelApiWac_RangeSort$invokeMethod$in(dispatchId, args) {
        switch (dispatchId) {
            case ExcelApiWac.RangeSort._dispiD_Apply$p:
                return this._applyInvoke$p$0(args);
            case ExcelApiWac.RangeSort._dispiD_ApplyAndReturnFirstField$p:
                return this._applyAndReturnFirstFieldInvoke$p$0(args);
            case ExcelApiWac.RangeSort._dispiD_ApplyMixed$p:
                return this._applyMixedInvoke$p$0(args);
            case ExcelApiWac.RangeSort._dispiD_ApplyQueryWithSortFieldAndReturnLast$p:
                return this._applyQueryWithSortFieldAndReturnLastInvoke$p$0(args);
        }
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokeMethodAsync: function ExcelApiWac_RangeSort$invokeMethodAsync$in(dispatchId, args) {
        var ret = this.invokeMethod(dispatchId, args);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(ret);
    },
    
    _applyInvoke$p$0: function ExcelApiWac_RangeSort$_applyInvoke$p$0$in(args) {
        var varFields = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'Array<ExcelApiWac.SortField>');
        return this.apply(varFields);
    },
    
    _applyAndReturnFirstFieldInvoke$p$0: function ExcelApiWac_RangeSort$_applyAndReturnFirstFieldInvoke$p$0$in(args) {
        var varFields = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'Array<ExcelApiWac.SortField>');
        return this.applyAndReturnFirstField(varFields);
    },
    
    _applyMixedInvoke$p$0: function ExcelApiWac_RangeSort$_applyMixedInvoke$p$0$in(args) {
        var varFields = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'Array<any>');
        return this.applyMixed(varFields);
    },
    
    _applyQueryWithSortFieldAndReturnLastInvoke$p$0: function ExcelApiWac_RangeSort$_applyQueryWithSortFieldAndReturnLastInvoke$p$0$in(args) {
        var varFields = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'Array<ExcelApiWac.QueryWithSortField>');
        return this.applyQueryWithSortFieldAndReturnLast(varFields);
    }
}


ExcelApiWac.SortField = function ExcelApiWac_SortField() {
}
ExcelApiWac.SortField.prototype = {
    
    get_assending: function ExcelApiWac_SortField$get_assending$in() {
        return false;
    },
    
    set_assending: function ExcelApiWac_SortField$set_assending$in(value) {
        return value;
    },
    
    get_columnIndex: function ExcelApiWac_SortField$get_columnIndex$in() {
        return 0;
    },
    
    set_columnIndex: function ExcelApiWac_SortField$set_columnIndex$in(value) {
        return value;
    },
    
    get_useCurrentCulture: function ExcelApiWac_SortField$get_useCurrentCulture$in() {
        return false;
    },
    
    set_useCurrentCulture: function ExcelApiWac_SortField$set_useCurrentCulture$in(value) {
        return value;
    },
    
    getTypeName: function ExcelApiWac_SortField$getTypeName$in() {
        return 'ExcelApi.SortField';
    },
    
    invokePropertyGet: function ExcelApiWac_SortField$invokePropertyGet$in(dispatchId) {
        switch (dispatchId) {
            case ExcelApiWac.SortField._dispiD_Assending$p:
                return this.get_assending();
            case ExcelApiWac.SortField._dispiD_ColumnIndex$p:
                return this.get_columnIndex();
            case ExcelApiWac.SortField._dispiD_UseCurrentCulture$p:
                return this.get_useCurrentCulture();
        }
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokePropertyGetAsync: function ExcelApiWac_SortField$invokePropertyGetAsync$in(dispatchId) {
        var ret = this.invokePropertyGet(dispatchId);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(ret);
    },
    
    invokePropertySet: function ExcelApiWac_SortField$invokePropertySet$in(dispatchId, args) {
        switch (dispatchId) {
            case ExcelApiWac.SortField._dispiD_Assending$p:
                this.set_assending(OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'boolean'));
                return;
            case ExcelApiWac.SortField._dispiD_ColumnIndex$p:
                this.set_columnIndex(OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'int'));
                return;
            case ExcelApiWac.SortField._dispiD_UseCurrentCulture$p:
                this.set_useCurrentCulture(OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'boolean'));
                return;
        }
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokePropertySetAsync: function ExcelApiWac_SortField$invokePropertySetAsync$in(dispatchId, args) {
        this.invokePropertySet(dispatchId, args);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(null);
    },
    
    invokeMethod: function ExcelApiWac_SortField$invokeMethod$in(dispatchId, args) {
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokeMethodAsync: function ExcelApiWac_SortField$invokeMethodAsync$in(dispatchId, args) {
        var ret = this.invokeMethod(dispatchId, args);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(ret);
    }
}


ExcelApiWac.TestCaseObject = function ExcelApiWac_TestCaseObject() {
}
ExcelApiWac.TestCaseObject.prototype = {
    
    calculateAddressAndSaveToRange: function ExcelApiWac_TestCaseObject$calculateAddressAndSaveToRange$in(street, city, range) {
        range.set_value(city);
        return street;
    },
    
    matrixSum: function ExcelApiWac_TestCaseObject$matrixSum$in(matrix) {
        return this._getDoubleValueFromObject$p$0(matrix);
    },
    
    sum: function ExcelApiWac_TestCaseObject$sum$in(values) {
        return this._getDoubleValueFromObject$p$0(values);
    },
    
    testParamBool: function ExcelApiWac_TestCaseObject$testParamBool$in(value) {
        return value;
    },
    
    testParamDouble: function ExcelApiWac_TestCaseObject$testParamDouble$in(value) {
        return value;
    },
    
    testParamFloat: function ExcelApiWac_TestCaseObject$testParamFloat$in(value) {
        return value;
    },
    
    testParamInt: function ExcelApiWac_TestCaseObject$testParamInt$in(value) {
        return value;
    },
    
    testParamRange: function ExcelApiWac_TestCaseObject$testParamRange$in(value) {
        return value;
    },
    
    testParamString: function ExcelApiWac_TestCaseObject$testParamString$in(value) {
        return value;
    },
    
    testUrlKeyValueDecode: function ExcelApiWac_TestCaseObject$testUrlKeyValueDecode$in(value) {
        return decodeURIComponent(value);
    },
    
    testUrlPathEncode: function ExcelApiWac_TestCaseObject$testUrlPathEncode$in(value) {
        return encodeURI(value);
    },
    
    getTypeName: function ExcelApiWac_TestCaseObject$getTypeName$in() {
        return 'ExcelApi.TestCaseObject';
    },
    
    invokePropertyGet: function ExcelApiWac_TestCaseObject$invokePropertyGet$in(dispatchId) {
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokePropertyGetAsync: function ExcelApiWac_TestCaseObject$invokePropertyGetAsync$in(dispatchId) {
        var ret = this.invokePropertyGet(dispatchId);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(ret);
    },
    
    invokePropertySet: function ExcelApiWac_TestCaseObject$invokePropertySet$in(dispatchId, args) {
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokePropertySetAsync: function ExcelApiWac_TestCaseObject$invokePropertySetAsync$in(dispatchId, args) {
        this.invokePropertySet(dispatchId, args);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(null);
    },
    
    invokeMethod: function ExcelApiWac_TestCaseObject$invokeMethod$in(dispatchId, args) {
        switch (dispatchId) {
            case ExcelApiWac.TestCaseObject._dispiD_CalculateAddressAndSaveToRange$p:
                return this._calculateAddressAndSaveToRangeInvoke$p$0(args);
            case ExcelApiWac.TestCaseObject._dispiD_MatrixSum$p:
                return this._matrixSumInvoke$p$0(args);
            case ExcelApiWac.TestCaseObject._dispiD_Sum$p:
                return this._sumInvoke$p$0(args);
            case ExcelApiWac.TestCaseObject._dispiD_TestParamBool$p:
                return this._testParamBoolInvoke$p$0(args);
            case ExcelApiWac.TestCaseObject._dispiD_TestParamDouble$p:
                return this._testParamDoubleInvoke$p$0(args);
            case ExcelApiWac.TestCaseObject._dispiD_TestParamFloat$p:
                return this._testParamFloatInvoke$p$0(args);
            case ExcelApiWac.TestCaseObject._dispiD_TestParamInt$p:
                return this._testParamIntInvoke$p$0(args);
            case ExcelApiWac.TestCaseObject._dispiD_TestParamRange$p:
                return this._testParamRangeInvoke$p$0(args);
            case ExcelApiWac.TestCaseObject._dispiD_TestParamString$p:
                return this._testParamStringInvoke$p$0(args);
            case ExcelApiWac.TestCaseObject._dispiD_TestUrlKeyValueDecode$p:
                return this._testUrlKeyValueDecodeInvoke$p$0(args);
            case ExcelApiWac.TestCaseObject._dispiD_TestUrlPathEncode$p:
                return this._testUrlPathEncodeInvoke$p$0(args);
        }
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokeMethodAsync: function ExcelApiWac_TestCaseObject$invokeMethodAsync$in(dispatchId, args) {
        var ret = this.invokeMethod(dispatchId, args);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(ret);
    },
    
    _calculateAddressAndSaveToRangeInvoke$p$0: function ExcelApiWac_TestCaseObject$_calculateAddressAndSaveToRangeInvoke$p$0$in(args) {
        var varStreet = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'string');
        var varCity = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 1, 'string');
        var varRange = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 2, 'ExcelApiWac.Range');
        return this.calculateAddressAndSaveToRange(varStreet, varCity, varRange);
    },
    
    _matrixSumInvoke$p$0: function ExcelApiWac_TestCaseObject$_matrixSumInvoke$p$0$in(args) {
        var varMatrix = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'Array<Array<any>>');
        return this.matrixSum(varMatrix);
    },
    
    _sumInvoke$p$0: function ExcelApiWac_TestCaseObject$_sumInvoke$p$0$in(args) {
        var varValues = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'Array<any>');
        return this.sum(varValues);
    },
    
    _testParamBoolInvoke$p$0: function ExcelApiWac_TestCaseObject$_testParamBoolInvoke$p$0$in(args) {
        var varValue = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'boolean');
        return this.testParamBool(varValue);
    },
    
    _testParamDoubleInvoke$p$0: function ExcelApiWac_TestCaseObject$_testParamDoubleInvoke$p$0$in(args) {
        var varValue = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'number');
        return this.testParamDouble(varValue);
    },
    
    _testParamFloatInvoke$p$0: function ExcelApiWac_TestCaseObject$_testParamFloatInvoke$p$0$in(args) {
        var varValue = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'number');
        return this.testParamFloat(varValue);
    },
    
    _testParamIntInvoke$p$0: function ExcelApiWac_TestCaseObject$_testParamIntInvoke$p$0$in(args) {
        var varValue = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'int');
        return this.testParamInt(varValue);
    },
    
    _testParamRangeInvoke$p$0: function ExcelApiWac_TestCaseObject$_testParamRangeInvoke$p$0$in(args) {
        var varValue = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'ExcelApiWac.Range');
        return this.testParamRange(varValue);
    },
    
    _testParamStringInvoke$p$0: function ExcelApiWac_TestCaseObject$_testParamStringInvoke$p$0$in(args) {
        var varValue = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'string');
        return this.testParamString(varValue);
    },
    
    _testUrlKeyValueDecodeInvoke$p$0: function ExcelApiWac_TestCaseObject$_testUrlKeyValueDecodeInvoke$p$0$in(args) {
        var varValue = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'string');
        return this.testUrlKeyValueDecode(varValue);
    },
    
    _testUrlPathEncodeInvoke$p$0: function ExcelApiWac_TestCaseObject$_testUrlPathEncodeInvoke$p$0$in(args) {
        var varValue = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'string');
        return this.testUrlPathEncode(varValue);
    },
    
    _getDoubleValueFromObject$p$0: function ExcelApiWac_TestCaseObject$_getDoubleValueFromObject$p$0$in(obj) {
        if (Array.isInstanceOfType(obj)) {
            var items = obj;
            var sum = 0;
            for (var i = 0; i < items.length; i++) {
                var item = this._getDoubleValueFromObject$p$0(items[i]);
                sum = sum + item;
            }
            return sum;
        }
        else if (ExcelApiWac.Range.isInstanceOfType(obj)) {
            var range = obj;
            return range.get_value();
        }
        return obj;
    }
}


ExcelApiWac.TestWorkbook = function ExcelApiWac_TestWorkbook() {
}
ExcelApiWac.TestWorkbook.prototype = {
    
    get_errorWorksheet: function ExcelApiWac_TestWorkbook$get_errorWorksheet$in() {
        throw OfficeExtension.WacRuntime.Utility.createRichApiException('Conflict', 'Conflict');
    },
    
    get_errorWorksheet2: function ExcelApiWac_TestWorkbook$get_errorWorksheet2$in() {
        throw OfficeExtension.WacRuntime.Utility.createRichApiException('Conflict2', 'Conflict');
    },
    
    errorMethod: function ExcelApiWac_TestWorkbook$errorMethod$in(input) {
        switch (input) {
            case 0:
                return 'NoError';
            case ExcelApiWac.ErrorMethodType.accessDenied:
                throw OfficeExtension.WacRuntime.Utility.createRichApiException('AccessDenied', 'Access Denied');
            case ExcelApiWac.ErrorMethodType.stateChanged:
                throw OfficeExtension.WacRuntime.Utility.createRichApiException('Conflict', 'Conflict');
            case ExcelApiWac.ErrorMethodType.bounds:
                throw OfficeExtension.WacRuntime.Utility.createRichApiException('OutOfRange', 'Out of range');
            case ExcelApiWac.ErrorMethodType.abort:
                throw OfficeExtension.WacRuntime.Utility.createRichApiException('Aborted', 'Aborted');
            default:
                throw OfficeExtension.WacRuntime.Utility.createRichApiException('Unexpected', 'Unexpected');
        }
    },
    
    errorMethod2: function ExcelApiWac_TestWorkbook$errorMethod2$in(input) {
        switch (input) {
            case 0:
                return 'NoError';
            case ExcelApiWac.ErrorMethodType.accessDenied:
                throw OfficeExtension.WacRuntime.Utility.createRichApiException('AccessDenied2', 'Access Denied');
            case ExcelApiWac.ErrorMethodType.stateChanged:
                throw OfficeExtension.WacRuntime.Utility.createRichApiException('Conflict', 'Conflict');
            case ExcelApiWac.ErrorMethodType.bounds:
                throw OfficeExtension.WacRuntime.Utility.createRichApiException('OutOfRange', 'Out of range');
            case ExcelApiWac.ErrorMethodType.abort:
                throw OfficeExtension.WacRuntime.Utility.createRichApiException('Aborted2', 'Aborted');
            default:
                throw OfficeExtension.WacRuntime.Utility.createRichApiException('Unexpected', 'Unexpected');
        }
    },
    
    getActiveWorksheet: function ExcelApiWac_TestWorkbook$getActiveWorksheet$in() {
        return new ExcelApiWac.Worksheet(0);
    },
    
    getCachedObjectCount: function ExcelApiWac_TestWorkbook$getCachedObjectCount$in() {
        return ExcelApiWac._rangeData.getCachedObjectCount();
    },
    
    getNullableBoolValue: function ExcelApiWac_TestWorkbook$getNullableBoolValue$in(nullable) {
        if (nullable) {
            return null;
        }
        return true;
    },
    
    getNullableEnumValue: function ExcelApiWac_TestWorkbook$getNullableEnumValue$in(nullable) {
        if (nullable) {
            return null;
        }
        return ExcelApiWac.ChartType.pie;
    },
    
    getObjectCount: function ExcelApiWac_TestWorkbook$getObjectCount$in() {
        return 0;
    },
    
    testNullableInputValue: function ExcelApiWac_TestWorkbook$testNullableInputValue$in(chartType, boolValue) {
        return null;
    },
    
    getTypeName: function ExcelApiWac_TestWorkbook$getTypeName$in() {
        return 'ExcelApi.TestWorkbook';
    },
    
    invokePropertyGet: function ExcelApiWac_TestWorkbook$invokePropertyGet$in(dispatchId) {
        switch (dispatchId) {
            case ExcelApiWac.TestWorkbook._dispiD_ErrorWorksheet$p:
                return this.get_errorWorksheet();
            case ExcelApiWac.TestWorkbook._dispiD_ErrorWorksheet2$p:
                return this.get_errorWorksheet2();
        }
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokePropertyGetAsync: function ExcelApiWac_TestWorkbook$invokePropertyGetAsync$in(dispatchId) {
        var ret = this.invokePropertyGet(dispatchId);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(ret);
    },
    
    invokePropertySet: function ExcelApiWac_TestWorkbook$invokePropertySet$in(dispatchId, args) {
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokePropertySetAsync: function ExcelApiWac_TestWorkbook$invokePropertySetAsync$in(dispatchId, args) {
        this.invokePropertySet(dispatchId, args);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(null);
    },
    
    invokeMethod: function ExcelApiWac_TestWorkbook$invokeMethod$in(dispatchId, args) {
        switch (dispatchId) {
            case ExcelApiWac.TestWorkbook._dispiD_ErrorMethod$p:
                return this._errorMethodInvoke$p$0(args);
            case ExcelApiWac.TestWorkbook._dispiD_ErrorMethod2$p:
                return this._errorMethod2Invoke$p$0(args);
            case ExcelApiWac.TestWorkbook._dispiD_GetActiveWorksheet$p:
                return this._getActiveWorksheetInvoke$p$0(args);
            case ExcelApiWac.TestWorkbook._dispiD_GetCachedObjectCount$p:
                return this._getCachedObjectCountInvoke$p$0(args);
            case ExcelApiWac.TestWorkbook._dispiD_GetNullableBoolValue$p:
                return this._getNullableBoolValueInvoke$p$0(args);
            case ExcelApiWac.TestWorkbook._dispiD_GetNullableEnumValue$p:
                return this._getNullableEnumValueInvoke$p$0(args);
            case ExcelApiWac.TestWorkbook._dispiD_GetObjectCount$p:
                return this._getObjectCountInvoke$p$0(args);
            case ExcelApiWac.TestWorkbook._dispiD_TestNullableInputValue$p:
                return this._testNullableInputValueInvoke$p$0(args);
        }
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokeMethodAsync: function ExcelApiWac_TestWorkbook$invokeMethodAsync$in(dispatchId, args) {
        var ret = this.invokeMethod(dispatchId, args);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(ret);
    },
    
    _errorMethodInvoke$p$0: function ExcelApiWac_TestWorkbook$_errorMethodInvoke$p$0$in(args) {
        var varInput = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'int');
        return this.errorMethod(varInput);
    },
    
    _errorMethod2Invoke$p$0: function ExcelApiWac_TestWorkbook$_errorMethod2Invoke$p$0$in(args) {
        var varInput = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'int');
        return this.errorMethod2(varInput);
    },
    
    _getActiveWorksheetInvoke$p$0: function ExcelApiWac_TestWorkbook$_getActiveWorksheetInvoke$p$0$in(args) {
        return this.getActiveWorksheet();
    },
    
    _getCachedObjectCountInvoke$p$0: function ExcelApiWac_TestWorkbook$_getCachedObjectCountInvoke$p$0$in(args) {
        return this.getCachedObjectCount();
    },
    
    _getNullableBoolValueInvoke$p$0: function ExcelApiWac_TestWorkbook$_getNullableBoolValueInvoke$p$0$in(args) {
        var varNullable = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'boolean');
        return this.getNullableBoolValue(varNullable);
    },
    
    _getNullableEnumValueInvoke$p$0: function ExcelApiWac_TestWorkbook$_getNullableEnumValueInvoke$p$0$in(args) {
        var varNullable = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'boolean');
        return this.getNullableEnumValue(varNullable);
    },
    
    _getObjectCountInvoke$p$0: function ExcelApiWac_TestWorkbook$_getObjectCountInvoke$p$0$in(args) {
        return this.getObjectCount();
    },
    
    _testNullableInputValueInvoke$p$0: function ExcelApiWac_TestWorkbook$_testNullableInputValueInvoke$p$0$in(args) {
        var varChartType = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'int?');
        var varBoolValue = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 1, 'boolean?');
        return this.testNullableInputValue(varChartType, varBoolValue);
    }
}


ExcelApiWac._utility = function ExcelApiWac__utility() {
}
ExcelApiWac._utility._createPromiseUsingTimeout$i = function ExcelApiWac__utility$_createPromiseUsingTimeout$i$st(value, milliseconds) {
    return OfficeExtension.WacRuntime.Utility.createPromise(function(resolve, reject) {
        window.setTimeout(function() {
            resolve(value);
        }, 0);
    });
}


ExcelApiWac.Workbook = function ExcelApiWac_Workbook() {
}
ExcelApiWac.Workbook.prototype = {
    
    getActiveWorksheet: function ExcelApiWac_Workbook$getActiveWorksheet$in() {
        return ExcelApiWac._utility._createPromiseUsingTimeout$i(new ExcelApiWac.Worksheet(0), 20);
    },
    
    get_charts: function ExcelApiWac_Workbook$get_charts$in() {
        return new ExcelApiWac.ChartCollection();
    },
    
    get_sheets: function ExcelApiWac_Workbook$get_sheets$in() {
        return new ExcelApiWac.WorksheetCollection();
    },
    
    getChartByType: function ExcelApiWac_Workbook$getChartByType$in(chartType) {
        return null;
    },
    
    getChartByTypeTitle: function ExcelApiWac_Workbook$getChartByTypeTitle$in(chartType, title) {
        return null;
    },
    
    someAction: function ExcelApiWac_Workbook$someAction$in(intVal, strVal, enumVal) {
        return null;
    },
    
    getTypeName: function ExcelApiWac_Workbook$getTypeName$in() {
        return 'ExcelApi.Workbook';
    },
    
    invokePropertyGet: function ExcelApiWac_Workbook$invokePropertyGet$in(dispatchId) {
        switch (dispatchId) {
            case ExcelApiWac.Workbook._dispiD_Charts$p:
                return this.get_charts();
            case ExcelApiWac.Workbook._dispiD_Sheets$p:
                return this.get_sheets();
        }
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokePropertyGetAsync: function ExcelApiWac_Workbook$invokePropertyGetAsync$in(dispatchId) {
        switch (dispatchId) {
            case ExcelApiWac.Workbook._dispiD_ActiveWorksheet$p:
                return this.getActiveWorksheet();
        }
        var ret = this.invokePropertyGet(dispatchId);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(ret);
    },
    
    invokePropertySet: function ExcelApiWac_Workbook$invokePropertySet$in(dispatchId, args) {
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokePropertySetAsync: function ExcelApiWac_Workbook$invokePropertySetAsync$in(dispatchId, args) {
        this.invokePropertySet(dispatchId, args);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(null);
    },
    
    invokeMethod: function ExcelApiWac_Workbook$invokeMethod$in(dispatchId, args) {
        switch (dispatchId) {
            case ExcelApiWac.Workbook._dispiD_GetChartByType$p:
                return this._getChartByTypeInvoke$p$0(args);
            case ExcelApiWac.Workbook._dispiD_GetChartByTypeTitle$p:
                return this._getChartByTypeTitleInvoke$p$0(args);
            case ExcelApiWac.Workbook._dispiD_SomeAction$p:
                return this._someActionInvoke$p$0(args);
        }
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokeMethodAsync: function ExcelApiWac_Workbook$invokeMethodAsync$in(dispatchId, args) {
        var ret = this.invokeMethod(dispatchId, args);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(ret);
    },
    
    _getChartByTypeInvoke$p$0: function ExcelApiWac_Workbook$_getChartByTypeInvoke$p$0$in(args) {
        var varChartType = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'int');
        return this.getChartByType(varChartType);
    },
    
    _getChartByTypeTitleInvoke$p$0: function ExcelApiWac_Workbook$_getChartByTypeTitleInvoke$p$0$in(args) {
        var varChartType = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'int');
        var varTitle = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 1, 'string');
        return this.getChartByTypeTitle(varChartType, varTitle);
    },
    
    _someActionInvoke$p$0: function ExcelApiWac_Workbook$_someActionInvoke$p$0$in(args) {
        var varIntVal = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'int');
        var varStrVal = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 1, 'string');
        var varEnumVal = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 2, 'int');
        return this.someAction(varIntVal, varStrVal, varEnumVal);
    }
}


ExcelApiWac.Worksheet = function ExcelApiWac_Worksheet(id) {
    this._m_id$p$0 = id;
}
ExcelApiWac.Worksheet.prototype = {
    _m_id$p$0: 0,
    _m_calculatedName$p$0: null,
    
    get_activeCell: function ExcelApiWac_Worksheet$get_activeCell$in() {
        return new ExcelApiWac.Range('A1');
    },
    
    get_activeCellInvalidAfterRequest: function ExcelApiWac_Worksheet$get_activeCellInvalidAfterRequest$in() {
        return new ExcelApiWac.Range('A1');
    },
    
    getCalculatedName: function ExcelApiWac_Worksheet$getCalculatedName$in() {
        if (!this._m_calculatedName$p$0) {
            this._m_calculatedName$p$0 = 'Sheet' + this._m_id$p$0;
        }
        return ExcelApiWac._utility._createPromiseUsingTimeout$i(this._m_calculatedName$p$0, 0);
    },
    
    setCalculatedName: function ExcelApiWac_Worksheet$setCalculatedName$in(value) {
        this._m_calculatedName$p$0 = value;
        return ExcelApiWac._utility._createPromiseUsingTimeout$i(null, 0);
    },
    
    get_name: function ExcelApiWac_Worksheet$get_name$in() {
        return 'Sheet' + this._m_id$p$0;
    },
    
    get__Id: function ExcelApiWac_Worksheet$get__Id$in() {
        return this._m_id$p$0;
    },
    
    nullChart: function ExcelApiWac_Worksheet$nullChart$in(address) {
        return null;
    },
    
    nullRange: function ExcelApiWac_Worksheet$nullRange$in(address) {
        return null;
    },
    
    getRange: function ExcelApiWac_Worksheet$getRange$in(address) {
        return ExcelApiWac._utility._createPromiseUsingTimeout$i(new ExcelApiWac.Range(address), 5);
    },
    
    someRangeOperation: function ExcelApiWac_Worksheet$someRangeOperation$in(input, range) {
        var $$t_5 = this;
        var p = OfficeExtension.WacRuntime.Utility.createPromise(function(resolve, reject) {
            window.setTimeout(function() {
                range.set_text(input);
                resolve(input);
            }, 15);
        });
        return p;
    },
    
    _RestOnly: function ExcelApiWac_Worksheet$_RestOnly$in() {
        return null;
    },
    
    getTypeName: function ExcelApiWac_Worksheet$getTypeName$in() {
        return 'ExcelApi.Worksheet';
    },
    
    invokePropertyGet: function ExcelApiWac_Worksheet$invokePropertyGet$in(dispatchId) {
        switch (dispatchId) {
            case ExcelApiWac.Worksheet._dispiD_ActiveCell$p:
                return this.get_activeCell();
            case ExcelApiWac.Worksheet._dispiD_ActiveCellInvalidAfterRequest$p:
                return this.get_activeCellInvalidAfterRequest();
            case ExcelApiWac.Worksheet._dispiD_Name$p:
                return this.get_name();
            case ExcelApiWac.Worksheet._dispiD__Id$p:
                return this.get__Id();
        }
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokePropertyGetAsync: function ExcelApiWac_Worksheet$invokePropertyGetAsync$in(dispatchId) {
        switch (dispatchId) {
            case ExcelApiWac.Worksheet._dispiD_CalculatedName$p:
                return this.getCalculatedName();
        }
        var ret = this.invokePropertyGet(dispatchId);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(ret);
    },
    
    invokePropertySet: function ExcelApiWac_Worksheet$invokePropertySet$in(dispatchId, args) {
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokePropertySetAsync: function ExcelApiWac_Worksheet$invokePropertySetAsync$in(dispatchId, args) {
        switch (dispatchId) {
            case ExcelApiWac.Worksheet._dispiD_CalculatedName$p:
                return this.setCalculatedName(OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'string'));
        }
        this.invokePropertySet(dispatchId, args);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(null);
    },
    
    invokeMethod: function ExcelApiWac_Worksheet$invokeMethod$in(dispatchId, args) {
        switch (dispatchId) {
            case ExcelApiWac.Worksheet._dispiD_NullChart$p:
                return this._nullChartInvoke$p$0(args);
            case ExcelApiWac.Worksheet._dispiD_NullRange$p:
                return this._nullRangeInvoke$p$0(args);
            case ExcelApiWac.Worksheet._dispiD__RestOnly$p:
                return this._RestOnlyInvoke$0(args);
        }
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokeMethodAsync: function ExcelApiWac_Worksheet$invokeMethodAsync$in(dispatchId, args) {
        switch (dispatchId) {
            case ExcelApiWac.Worksheet._dispiD_Range$p:
                return this._rangeInvoke$p$0(args);
            case ExcelApiWac.Worksheet._dispiD_SomeRangeOperation$p:
                return this._someRangeOperationInvoke$p$0(args);
        }
        var ret = this.invokeMethod(dispatchId, args);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(ret);
    },
    
    _nullChartInvoke$p$0: function ExcelApiWac_Worksheet$_nullChartInvoke$p$0$in(args) {
        var varAddress = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'string');
        return this.nullChart(varAddress);
    },
    
    _nullRangeInvoke$p$0: function ExcelApiWac_Worksheet$_nullRangeInvoke$p$0$in(args) {
        var varAddress = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'string');
        return this.nullRange(varAddress);
    },
    
    _rangeInvoke$p$0: function ExcelApiWac_Worksheet$_rangeInvoke$p$0$in(args) {
        var varAddress = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'string');
        return this.getRange(varAddress);
    },
    
    _someRangeOperationInvoke$p$0: function ExcelApiWac_Worksheet$_someRangeOperationInvoke$p$0$in(args) {
        var varInput = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'string');
        var varRange = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 1, 'ExcelApiWac.Range');
        return this.someRangeOperation(varInput, varRange);
    },
    
    _RestOnlyInvoke$0: function ExcelApiWac_Worksheet$_RestOnlyInvoke$0$in(args) {
        return this._RestOnly();
    }
}


ExcelApiWac.WorksheetCollection = function ExcelApiWac_WorksheetCollection() {
}
ExcelApiWac.WorksheetCollection.prototype = {
    
    add: function ExcelApiWac_WorksheetCollection$add$in(name) {
        return new ExcelApiWac.Worksheet(2006);
    },
    
    findSheet: function ExcelApiWac_WorksheetCollection$findSheet$in(text) {
        return ExcelApiWac._utility._createPromiseUsingTimeout$i(new ExcelApiWac.Worksheet(2004), 10);
    },
    
    getActiveWorksheetInvalidAfterRequest: function ExcelApiWac_WorksheetCollection$getActiveWorksheetInvalidAfterRequest$in() {
        return new ExcelApiWac.Worksheet(2004);
    },
    
    getItem: function ExcelApiWac_WorksheetCollection$getItem$in(index) {
        return this._GetItem(index);
    },
    
    _GetItem: function ExcelApiWac_WorksheetCollection$_GetItem$in(index) {
        if (typeof(index) === 'number') {
            var id = index;
            return new ExcelApiWac.Worksheet(id);
        }
        else if (typeof(index) === 'string') {
            var id = this._getSheetIdFromName$p$0(index);
            return new ExcelApiWac.Worksheet(id);
        }
        else {
            throw Error.argument();
        }
    },
    
    getAsyncEnumerator: function ExcelApiWac_WorksheetCollection$getAsyncEnumerator$in(skip, top) {
        return ExcelApiWac._utility._createPromiseUsingTimeout$i(new ExcelApiWac._worksheetEnumerator(skip), 0);
    },
    
    getTypeName: function ExcelApiWac_WorksheetCollection$getTypeName$in() {
        return 'ExcelApi.WorksheetCollection';
    },
    
    invokePropertyGet: function ExcelApiWac_WorksheetCollection$invokePropertyGet$in(dispatchId) {
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokePropertyGetAsync: function ExcelApiWac_WorksheetCollection$invokePropertyGetAsync$in(dispatchId) {
        var ret = this.invokePropertyGet(dispatchId);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(ret);
    },
    
    invokePropertySet: function ExcelApiWac_WorksheetCollection$invokePropertySet$in(dispatchId, args) {
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokePropertySetAsync: function ExcelApiWac_WorksheetCollection$invokePropertySetAsync$in(dispatchId, args) {
        this.invokePropertySet(dispatchId, args);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(null);
    },
    
    invokeMethod: function ExcelApiWac_WorksheetCollection$invokeMethod$in(dispatchId, args) {
        switch (dispatchId) {
            case ExcelApiWac.WorksheetCollection._dispiD_Add$p:
                return this._addInvoke$p$0(args);
            case ExcelApiWac.WorksheetCollection._dispiD_GetActiveWorksheetInvalidAfterRequest$p:
                return this._getActiveWorksheetInvalidAfterRequestInvoke$p$0(args);
            case ExcelApiWac.WorksheetCollection._dispiD_GetItem$p:
                return this._getItemInvoke$p$0(args);
            case ExcelApiWac.WorksheetCollection._dispiD_Item$p:
                return this._GetItemInvoke$0(args);
        }
        throw OfficeExtension.WacRuntime.Utility.createApiNotFoundException(dispatchId);
    },
    
    invokeMethodAsync: function ExcelApiWac_WorksheetCollection$invokeMethodAsync$in(dispatchId, args) {
        switch (dispatchId) {
            case ExcelApiWac.WorksheetCollection._dispiD_FindSheet$p:
                return this._findSheetInvoke$p$0(args);
        }
        var ret = this.invokeMethod(dispatchId, args);
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(ret);
    },
    
    _addInvoke$p$0: function ExcelApiWac_WorksheetCollection$_addInvoke$p$0$in(args) {
        var varName = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'string');
        return this.add(varName);
    },
    
    _findSheetInvoke$p$0: function ExcelApiWac_WorksheetCollection$_findSheetInvoke$p$0$in(args) {
        var varText = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'string');
        return this.findSheet(varText);
    },
    
    _getActiveWorksheetInvalidAfterRequestInvoke$p$0: function ExcelApiWac_WorksheetCollection$_getActiveWorksheetInvalidAfterRequestInvoke$p$0$in(args) {
        return this.getActiveWorksheetInvalidAfterRequest();
    },
    
    _getItemInvoke$p$0: function ExcelApiWac_WorksheetCollection$_getItemInvoke$p$0$in(args) {
        var varIndex = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'any');
        return this.getItem(varIndex);
    },
    
    _GetItemInvoke$0: function ExcelApiWac_WorksheetCollection$_GetItemInvoke$0$in(args) {
        var varIndex = OfficeExtension.WacRuntime.Utility.getAndValidateArgument(args, 0, 'any');
        return this._GetItem(varIndex);
    },
    
    _getSheetIdFromName$p$0: function ExcelApiWac_WorksheetCollection$_getSheetIdFromName$p$0$in(name) {
        if (name.substr(0, 'Sheet'.length) === 'Sheet') {
            var id = name.substr('Sheet'.length);
            return parseInt(id);
        }
        throw Error.argument();
    }
}


ExcelApiWac._worksheetEnumerator = function ExcelApiWac__worksheetEnumerator(skip) {
    this._m_index$p$0 = -1;
    if (skip > 0) {
        this._m_index$p$0 += skip;
    }
}
ExcelApiWac._worksheetEnumerator.prototype = {
    
    get_current: function ExcelApiWac__worksheetEnumerator$get_current$in() {
        return new ExcelApiWac.Worksheet(this._m_index$p$0 + 2000);
    },
    
    moveNext: function ExcelApiWac__worksheetEnumerator$moveNext$in() {
        var $$t_4 = this;
        return OfficeExtension.WacRuntime.Utility.createPromise(function(resolve, reject) {
            window.setTimeout(function() {
                $$t_4._m_index$p$0++;
                if ($$t_4._m_index$p$0 < 6) {
                    resolve(true);
                }
                else {
                    resolve(false);
                }
            }, 0);
        });
    }
}


Type.registerNamespace('FakeXlapiWac');

FakeXlapiWac.Tests = function FakeXlapiWac_Tests() {
}
FakeXlapiWac.Tests.testPromiseWac = function FakeXlapiWac_Tests$testPromiseWac$st() {
    var success = false;
    OfficeExtension.WacRuntime.Utility.createPromiseFromResult('abc').then(function(result) {
        return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(result.toString() + 'def');
    }).then(function(result) {
        if (result.toString() === 'abcdef') {
            success = true;
        }
        RichApiTest.log.done(success);;
        return null;
    });
}
FakeXlapiWac.Tests.testExecuteRichApi = function FakeXlapiWac_Tests$testExecuteRichApi$st() {
    var app = new ExcelApiWac.Application();
    var strRequest = '[\"IterativeExecutor\",\"POST\",\"ProcessQuery\",[],\"{\\\"Actions\\\":[{\\\"Id\\\":2,\\\"ActionType\\\":1,\\\"Name\\\":\\\"\\\",\\\"ObjectPathId\\\":1},{\\\"Id\\\":4,\\\"ActionType\\\":1,\\\"Name\\\":\\\"\\\",\\\"ObjectPathId\\\":3},{\\\"Id\\\":6,\\\"ActionType\\\":1,\\\"Name\\\":\\\"\\\",\\\"ObjectPathId\\\":5},{\\\"Id\\\":8,\\\"ActionType\\\":1,\\\"Name\\\":\\\"\\\",\\\"ObjectPathId\\\":7},{\\\"Id\\\":9,\\\"ActionType\\\":4,\\\"Name\\\":\\\"Text\\\",\\\"ObjectPathId\\\":7,\\\"ArgumentInfo\\\":{\\\"Arguments\\\":[\\\"abc\\\"]}},{\\\"Id\\\":10,\\\"ActionType\\\":2,\\\"Name\\\":\\\"\\\",\\\"ObjectPathId\\\":7,\\\"QueryInfo\\\":{\\\"Select\\\":[\\\"Text\\\"]}}],\\\"ObjectPaths\\\":{\\\"1\\\":{\\\"Id\\\":1,\\\"ObjectPathType\\\":1,\\\"Name\\\":\\\"\\\"},\\\"3\\\":{\\\"Id\\\":3,\\\"ObjectPathType\\\":4,\\\"Name\\\":\\\"ActiveWorkbook\\\",\\\"ParentObjectPathId\\\":1},\\\"5\\\":{\\\"Id\\\":5,\\\"ObjectPathType\\\":4,\\\"Name\\\":\\\"ActiveWorksheet\\\",\\\"ParentObjectPathId\\\":3},\\\"7\\\":{\\\"Id\\\":7,\\\"ObjectPathType\\\":3,\\\"Name\\\":\\\"Range\\\",\\\"ParentObjectPathId\\\":5,\\\"ArgumentInfo\\\":{\\\"Arguments\\\":[\\\"A1\\\"]}}}}\",0,1,\"\",\"\",\"\"]';
    var request = JSON.parse(strRequest);
    var success = false;
    OfficeExtension.WacRuntime.RichApi.executeRichApiRequest(null, ExcelApiWac.TypeRegister.registerTypes, app, request).then(function(result) {
        var str = JSON.stringify(result);
        RichApiTest.log.comment(str);
        success = true;
        RichApiTest.log.done(success);;
        return null;
    });
}


ExcelApiWac.Application.registerClass('ExcelApiWac.Application');
ExcelApiWac.Chart.registerClass('ExcelApiWac.Chart');
ExcelApiWac.ChartCollection.registerClass('ExcelApiWac.ChartCollection');
ExcelApiWac.TypeRegister.registerClass('ExcelApiWac.TypeRegister');
ExcelApiWac.QueryWithSortField.registerClass('ExcelApiWac.QueryWithSortField');
ExcelApiWac.Range.registerClass('ExcelApiWac.Range');
ExcelApiWac._rangeData.registerClass('ExcelApiWac._rangeData');
ExcelApiWac.RangeSort.registerClass('ExcelApiWac.RangeSort');
ExcelApiWac.SortField.registerClass('ExcelApiWac.SortField');
ExcelApiWac.TestCaseObject.registerClass('ExcelApiWac.TestCaseObject');
ExcelApiWac.TestWorkbook.registerClass('ExcelApiWac.TestWorkbook');
ExcelApiWac._utility.registerClass('ExcelApiWac._utility');
ExcelApiWac.Workbook.registerClass('ExcelApiWac.Workbook');
ExcelApiWac.Worksheet.registerClass('ExcelApiWac.Worksheet');
ExcelApiWac.WorksheetCollection.registerClass('ExcelApiWac.WorksheetCollection');
ExcelApiWac._worksheetEnumerator.registerClass('ExcelApiWac._worksheetEnumerator', null, OfficeExtension.WacRuntime.IAsyncEnumerator);
FakeXlapiWac.Tests.registerClass('FakeXlapiWac.Tests');
ExcelApiWac.Application._dispiD_ActiveWorkbook$p = 1;
ExcelApiWac.Application._dispiD_TestWorkbook$p = 2;
ExcelApiWac.Application._dispiD__GetObjectTypeNameByReferenceId$p = 3;
ExcelApiWac.Application._dispiD__GetObjectByReferenceId$p = 4;
ExcelApiWac.Application._dispiD__RemoveReference$p = 5;
ExcelApiWac.Application._dispiD_HasBase$p = 6;
ExcelApiWac.Chart._dispiD_Name$p = 1;
ExcelApiWac.Chart._dispiD_ChartType$p = 2;
ExcelApiWac.Chart._dispiD_Title$p = 3;
ExcelApiWac.Chart._dispiD_Delete$p = 4;
ExcelApiWac.Chart._dispiD_Id$p = 5;
ExcelApiWac.Chart._dispiD_NullableChartType$p = 6;
ExcelApiWac.Chart._dispiD_NullableShowLabel$p = 7;
ExcelApiWac.Chart._dispiD_ImageData$p = 8;
ExcelApiWac.Chart._dispiD_GetAsImage$p = 9;
ExcelApiWac.ChartCollection._dispiD_Count$p = 1;
ExcelApiWac.ChartCollection._dispiD_Item$p = 2;
ExcelApiWac.ChartCollection._dispiD_Add$p = 3;
ExcelApiWac.ChartCollection._dispiD_GetItemAt$p = 4;
ExcelApiWac.QueryWithSortField._dispiD_RowLimit$p = 1;
ExcelApiWac.QueryWithSortField._dispiD_Field$p = 2;
ExcelApiWac.Range._dispiD_RowIndex$p = 1;
ExcelApiWac.Range._dispiD_ColumnIndex$p = 2;
ExcelApiWac.Range._dispiD_Value$p = 3;
ExcelApiWac.Range._dispiD_ReplaceValue$p = 4;
ExcelApiWac.Range._dispiD_Activate$p = 5;
ExcelApiWac.Range._dispiD_Text$p = 6;
ExcelApiWac.Range._dispiD_ValueArray$p = 7;
ExcelApiWac.Range._dispiD_TextArray$p = 8;
ExcelApiWac.Range._dispiD_LogText$p = 9;
ExcelApiWac.Range._dispiD__ReferenceId$p = 10;
ExcelApiWac.Range._dispiD__KeepReference$p = 11;
ExcelApiWac.Range._dispiD_ValueArray2$p = 12;
ExcelApiWac.Range._dispiD_GetValueArray2$p = 13;
ExcelApiWac.Range._dispiD_SetValueArray2$p = 14;
ExcelApiWac.Range._dispiD__OnAccess$p = 15;
ExcelApiWac.Range._dispiD_ValueTypes$p = 16;
ExcelApiWac.Range._dispiD_NotRestMethod$p = 17;
ExcelApiWac.Range._dispiD_Sort$p = 18;
ExcelApiWac._rangeData._s_cache$p = null;
ExcelApiWac._rangeData._s_index$p = 0;
ExcelApiWac.RangeSort._dispiD_Apply$p = 1;
ExcelApiWac.RangeSort._dispiD_Fields$p = 2;
ExcelApiWac.RangeSort._dispiD_ApplyAndReturnFirstField$p = 3;
ExcelApiWac.RangeSort._dispiD_QueryField$p = 4;
ExcelApiWac.RangeSort._dispiD_ApplyQueryWithSortFieldAndReturnLast$p = 5;
ExcelApiWac.RangeSort._dispiD_ApplyMixed$p = 6;
ExcelApiWac.RangeSort._dispiD_Fields2$p = 7;
ExcelApiWac.SortField._dispiD_ColumnIndex$p = 1;
ExcelApiWac.SortField._dispiD_Assending$p = 2;
ExcelApiWac.SortField._dispiD_UseCurrentCulture$p = 3;
ExcelApiWac.TestCaseObject._dispiD_CalculateAddressAndSaveToRange$p = 1;
ExcelApiWac.TestCaseObject._dispiD_TestParamBool$p = 2;
ExcelApiWac.TestCaseObject._dispiD_TestParamInt$p = 3;
ExcelApiWac.TestCaseObject._dispiD_TestParamFloat$p = 4;
ExcelApiWac.TestCaseObject._dispiD_TestParamDouble$p = 5;
ExcelApiWac.TestCaseObject._dispiD_TestParamString$p = 6;
ExcelApiWac.TestCaseObject._dispiD_TestParamRange$p = 7;
ExcelApiWac.TestCaseObject._dispiD_TestUrlKeyValueDecode$p = 8;
ExcelApiWac.TestCaseObject._dispiD_TestUrlPathEncode$p = 9;
ExcelApiWac.TestCaseObject._dispiD_Sum$p = 10;
ExcelApiWac.TestCaseObject._dispiD_MatrixSum$p = 11;
ExcelApiWac.TestWorkbook._dispiD_ErrorWorksheet$p = 1;
ExcelApiWac.TestWorkbook._dispiD_ErrorMethod$p = 2;
ExcelApiWac.TestWorkbook._dispiD_GetObjectCount$p = 3;
ExcelApiWac.TestWorkbook._dispiD_GetActiveWorksheet$p = 4;
ExcelApiWac.TestWorkbook._dispiD_ErrorWorksheet2$p = 5;
ExcelApiWac.TestWorkbook._dispiD_ErrorMethod2$p = 6;
ExcelApiWac.TestWorkbook._dispiD_TestNullableInputValue$p = 7;
ExcelApiWac.TestWorkbook._dispiD_GetNullableEnumValue$p = 8;
ExcelApiWac.TestWorkbook._dispiD_GetNullableBoolValue$p = 9;
ExcelApiWac.TestWorkbook._dispiD_GetCachedObjectCount$p = 10;
ExcelApiWac.Workbook._dispiD_Sheets$p = 1;
ExcelApiWac.Workbook._dispiD_ActiveWorksheet$p = 2;
ExcelApiWac.Workbook._dispiD_Charts$p = 3;
ExcelApiWac.Workbook._dispiD_GetChartByType$p = 4;
ExcelApiWac.Workbook._dispiD_GetChartByTypeTitle$p = 5;
ExcelApiWac.Workbook._dispiD_SomeAction$p = 6;
ExcelApiWac.Worksheet._dispiD_Range$p = 1;
ExcelApiWac.Worksheet._dispiD_SomeRangeOperation$p = 2;
ExcelApiWac.Worksheet._dispiD_Name$p = 3;
ExcelApiWac.Worksheet._dispiD_ActiveCell$p = 4;
ExcelApiWac.Worksheet._dispiD_ActiveCellInvalidAfterRequest$p = 5;
ExcelApiWac.Worksheet._dispiD__Id$p = 6;
ExcelApiWac.Worksheet._dispiD_CalculatedName$p = 7;
ExcelApiWac.Worksheet._dispiD_NullRange$p = 8;
ExcelApiWac.Worksheet._dispiD_NullChart$p = 9;
ExcelApiWac.Worksheet._dispiD__RestOnly$p = 10;
ExcelApiWac.WorksheetCollection._dispiD_Item$p = 1;
ExcelApiWac.WorksheetCollection._dispiD_GetActiveWorksheetInvalidAfterRequest$p = 2;
ExcelApiWac.WorksheetCollection._dispiD_GetItem$p = 3;
ExcelApiWac.WorksheetCollection._dispiD_FindSheet$p = 4;
ExcelApiWac.WorksheetCollection._dispiD_Add$p = 5;
