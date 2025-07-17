# AltaworxDeviceBulkChange Lambda - ICCID/IMEI Processing Analysis

## Overview
The `AltaworxDeviceBulkChange` lambda function is triggered via AWS SQS when bulk device changes are processed. The ICCID/IMEI handling occurs primarily in the `MobilityController.cs` file.

## Key Methods and Flow

### 1. Entry Point - PostChangeICCIDorIMEI
**Location**: `MobilityController.cs` lines 511-543
```csharp
[HttpPost]
public async Task<ActionResult> PostChangeICCIDorIMEI(BulkchangeUpdateICCIDorIMEI model)
{
    // ... validation code ...
    var changeType = DeviceChangeType.ChangeICCIDorIMEI;
    var bulkChange = new DeviceBulkChange
    {
        ChangeRequestTypeId = (int)changeType,
        ServiceProviderId = model.ServiceProviderId,
        // ... other properties ...
        Mobility_DeviceChange = BuildUpdateICCIDorIMEI(altaWrxDb, Session, permissionManager, model).ToList()
    };
    var bulkChangeId = changeRepository.CreateBulkChange(bulkChange);
    await ProcessBulkChange(bulkChange.id);
    // ...
}
```

### 2. ICCID/IMEI Processing Logic - BuildUpdateICCIDorIMEI
**Location**: `MobilityController.cs` lines 3614-3653
```csharp
internal static IEnumerable<Mobility_DeviceChange> BuildUpdateICCIDorIMEI(
    AltaWorxCentral_Entities awxDb,
    HttpSessionStateBase session, 
    PermissionManager permissionManager, 
    BulkchangeUpdateICCIDorIMEI model)
{
    var createdBy = SessionHelper.GetAuditByName(session);
    var devicesByPhoneNumbers = GetDevicesByNumber(awxDb, model.ServiceProviderId, model.Devices);
    var archivedMSISDNs = CheckDevicesArchived(awxDb, model.ServiceProviderId, model.Devices);
    var deviceChanges = new List<Mobility_DeviceChange>();
    
    foreach (var modelDevice in model.Devices.Select((item, index) => new { item, index }))
    {
        var phoneNumber = modelDevice.item;
        if (!string.IsNullOrWhiteSpace(phoneNumber))
        {
            // Device validation logic
            if (!devicesByPhoneNumbers.ContainsKey(phoneNumber) || !devicesByPhoneNumbers.TryGetValue(phoneNumber, out var device))
            {
                // Error handling for missing devices
            }
            else
            {
                var newICCID = string.Empty;
                if (model.NewICCIDs != null && model.NewICCIDs.Count > modelDevice.index)
                {
                    newICCID = model.NewICCIDs[modelDevice.index];
                }
                var newIMEI = string.Empty;
                if (model.NewIMEIs != null && model.NewIMEIs.Count > modelDevice.index)
                {
                    newIMEI = model.NewIMEIs[modelDevice.index];
                }

                // **THIS IS WHERE YOUR CODE UPDATE SHOULD GO**
                deviceChanges.Add(new Mobility_DeviceChange(
                    CreateUpdateICCIDorIMEIChangeRequest(newICCID, newIMEI, device, device.ServiceZipCode, CommonStrings.UpdateICCIDorIMEIReasonCode, device.TechnologyType, model), 
                    device.id, 
                    device.ICCID, 
                    phoneNumber, 
                    createdBy));
            }
        }
    }
    return deviceChanges;
}
```

### 3. Change Request Creation - CreateUpdateICCIDorIMEIChangeRequest
**Location**: `MobilityController.cs` lines 3657-3700+
```csharp
private static string CreateUpdateICCIDorIMEIChangeRequest(string iccid, string imei, MobilityDevice device, string zipCode, string reasonCode, string technologyType, BulkchangeUpdateICCIDorIMEI model)
{
    var characteristicList = new List<ServiceCharacteristic>
    {
        new ServiceCharacteristic
        {
            Name = "reasonCode",
            Value = reasonCode
        },
        new ServiceCharacteristic
        {
            Name = "serviceZipCode",
            Value = zipCode
        },
        new ServiceCharacteristic
        {
            Name = "technologyType",
            Value = string.IsNullOrEmpty(technologyType) ? Resources.CommonStrings.MobiltiyTechnologyTypeDefault : technologyType
        },
        new ServiceCharacteristic
        {
            Name = "IMEI",
            Value = string.IsNullOrEmpty(imei) ? device.IMEI : imei
        },
        new ServiceCharacteristic
        {
            Name = "sim",
            Value = string.IsNullOrEmpty(iccid) ? device.ICCID : iccid
        }
    };
    
    // **THIS IS WHERE YOUR ICCID/IMEI LOGIC GETS BUILT INTO THE REQUEST**
    var changeEquipmentRequest = new TelegenceUpdateICCIDorIMEIRequest()
    {
        ServiceCharacteristic = characteristicList
    };
    var request = new BulkChangeStatusUpdateRequest<TelegenceUpdateICCIDorIMEIRequest>
    {
        Request = changeEquipmentRequest
    };
    
    // Additional processing for customer rate plan changes...
}
```

### 4. SQS Processing - ProcessBulkChange
**Location**: `MobilityController.cs` lines 2348-2380
```csharp
public async Task<ActionResult> ProcessBulkChange(long id, long additionBulkChangeId = 0)
{
    // ... validation ...
    if (bulkChange.Mobility_DeviceChange.Any(change => !change.IsProcessed))
    {
        newBulkChangeStatus = BulkChangeStatus.PROCESSING;
        var customObjectDbList = GetTenantCustomFields();
        var awsAccessKey = AwsAccessKeyFromCustomObjects(customObjectDbList);
        var awsSecretAccessKey = AwsSecretAccessKeyFromCustomObjects(customObjectDbList);
        var queueName = ValueFromCustomObjects(customObjectDbList, CommonConstants.CUSTOM_OBJECT_BULK_CHANGE_QUEUE_KEY);
        var sqsHelper = new SqsHelper(awsAccessKey, awsSecretAccessKey);
        
        // **THIS SENDS THE MESSAGE TO SQS WHICH TRIGGERS THE LAMBDA**
        var errorMessage = await sqsHelper.EnqueueBulkChangeAsync(queueName, id, additionBulkChangeId);
    }
}
```

### 5. Lambda Processing Indication
**Location**: `MobilityController.cs` line 3445
```csharp
ProcessBy = "AltaworxDeviceBulkChange",  // This indicates lambda processing
```

## Where to Add Your Updated Code

### Option 1: Modify BuildUpdateICCIDorIMEI Method
Add your ICCID/IMEI validation/processing logic in the `BuildUpdateICCIDorIMEI` method around lines 3639-3650 where the new ICCID and IMEI values are being processed.

### Option 2: Modify CreateUpdateICCIDorIMEIChangeRequest Method  
Add your logic in the `CreateUpdateICCIDorIMEIChangeRequest` method around lines 3670-3685 where the ServiceCharacteristic list is being built with IMEI and sim (ICCID) values.

### Option 3: Add Additional Validation
You could add validation methods that are called before the device changes are created, similar to how `CheckDevicesArchived` is used.

## Recommendation
I recommend updating the `CreateUpdateICCIDorIMEIChangeRequest` method to include your additional ICCID/IMEI processing logic, as this is where the final request structure is built that gets sent to the Lambda function via SQS.

Please share your specific code requirements so I can provide the exact updated code implementation.