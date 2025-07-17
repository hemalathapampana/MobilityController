using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using System.Web.ModelBinding;
using System.Web.Mvc;
using Amazon.SQS;
using Amazon.SQS.Model;
using Amop.Core.Constants;
using Amop.Core.Enumerations;
using Amop.Core.Models;
using Amop.Core.Models.DeviceBulkChange;
using Amop.Core.Models.Revio;
using Amop.Core.Models.Telegence;
using Amop.Core.Models.Telegence.Api;
using Amop.Core.Services.Http;
using Amop.Core.Services.Revio;
using Amop.Core.Services.Telegence;
using AttEBondingService;
using CsvHelper;
using CsvHelper.Configuration;
using KeySys.BaseMultiTenant.Helpers;
using KeySys.BaseMultiTenant.Mapping;
using KeySys.BaseMultiTenant.Models;
using KeySys.BaseMultiTenant.Models.BulkChange;
using KeySys.BaseMultiTenant.Models.CustomClasses;
using KeySys.BaseMultiTenant.Models.CustomerRatePool;
using KeySys.BaseMultiTenant.Models.Device;
using KeySys.BaseMultiTenant.Models.Mobility;
using KeySys.BaseMultiTenant.Models.Repositories;
using KeySys.BaseMultiTenant.Models.RevCustomer;
using KeySys.BaseMultiTenant.Models.Telegence;
using KeySys.BaseMultiTenant.Repositories;
using KeySys.BaseMultiTenant.Repositories.BillingPeriod;
using KeySys.BaseMultiTenant.Repositories.CarrierRatePlan;
using KeySys.BaseMultiTenant.Repositories.CustomerRatePool;
using KeySys.BaseMultiTenant.Repositories.Device;
using KeySys.BaseMultiTenant.Repositories.Mobility;
using KeySys.BaseMultiTenant.Repositories.Rev;
using KeySys.BaseMultiTenant.Repositories.Site;
using KeySys.BaseMultiTenant.Repositories.SiteGroup;
using KeySys.BaseMultiTenant.Resources;
using KeySys.BaseMultiTenant.Utilities;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using static KeySys.BaseMultiTenant.Helpers.Utils;
using MobilityDevice = KeySys.BaseMultiTenant.Models.Repositories.MobilityDevice;
using Service = Amop.Core.Models.Telegence.Api.Service;
using FileSystem = System.IO.File;
using KeySys.BaseMultiTenant.Controllers.AmopInternal;

namespace KeySys.BaseMultiTenant.Controllers
{
    public class MobilityController : AmopBaseController
    {
        private const string MOBILITY_LINE_CONFIGURATION_QUEUE_NAME = "Mobility Line Configuration Queue";
        private const PortalTypes PORTAL_TYPE = PortalTypes.Mobility;
        private const int ARCHIVAL_RECENT_USAGE_VALIDATION_DAYS = 30;
        private const string RELATED_PARTY_ROLE = "subscriber";
        private readonly IHttpClientFactory httpClientFactory;
        private readonly IHttpRequestFactory httpRequestFactory;
        private const string REMOVE_SOCCODE_STRING = "removeOfferingCode";
        private const int NUMBER_OF_EXCEL_SHEETS = 2;
        private const int MAXIMUM_CHARACTER_FILE_NAME = 255;
        private const int POSITION_OF_WORKSHEET_TWO = 2;

        // GET: Mobility
        public ActionResult Index([QueryString] AdvancedFilter advancedFilter = null, string filter = "", int page = 1, string sort = "", string sortDir = "")
        {
            ViewBag.PageTitle = "Mobility Inventory";

            if (permissionManager.UserCannotAccess(Session, ModuleEnum.Mobility))
                return RedirectToAction("Index", "Home");

            var repository = new MobilityInventoryRepository(altaWrxDb, permissionManager);
            var model = new MobilityInventoryModel
            {
                AdvancedFilter = advancedFilter?.WithMultiselectDeserializationFix(),
                Filter = filter,
                UsageAggregates = altaWrxDb.MobilityDeviceUsageAggregates.Where(x => x.IsActive && !x.IsDeleted).ToList(),
                DeviceInventoryList = PagedList.ToPagedList(repository.GetInventory(filter, advancedFilter, page, sort, sortDir), repository.GetInventoryCount(filter, advancedFilter)),
                AdvancedFilterServiceProviders = AdvancedFilterServiceProviders(),
                AdvancedFilterStatuses = AdvancedFilterStatuses(),
                TelegenceAllowedStatusList = ListHelper.TelegenceAllowedStatusList(altaWrxDb, permissionManager),
                eBondingStatusUpdateModalModel = new eBondingStatusUpdateModalModel
                {
                    AllowedStatusList = ListHelper.DeviceStatuses(altaWrxDb, IntegrationType.eBonding)
                        .Where(status => status.AllowsApiUpdate)
                        .Select(status => new eBondingStatus
                        {
                            Id = status.id,
                            DisplayName = status.DisplayName,
                            Status = status.Status,
                            ValidFrom = eBondingHelper.GetEbondingValidFromStatus(status.Status)
                        }),
                    States = ListHelper.States()
                }
            };

            return View(model);
        }

        public FileContentResult MobilityInventoryExport([QueryString] AdvancedFilter advancedFilter = null, string filter = "")
        {
            if (!permissionManager.UserCanAccessPortalTypeModule(Session, PORTAL_TYPE))
            {
                return null;
            }

            var repository = new MobilityInventoryRepository(altaWrxDb, permissionManager);
            var lineItems = repository.GetInventory(filter, advancedFilter).Select(device => device.ToMobilityInventoryExport()).ToList();

            var data = CreateDataSetForExport(lineItems);

            var bytes = ExcelUtilities.Export(data);

            return File(bytes, ExcelContentType, $"ReportMobilityInventory_{FileNameTimestamp()}.{ExcelFileExtension}");
        }

        private DataSet CreateDataSetForExport(IEnumerable<MobilityInventoryExport> lineItems)
        {
            var tableName = "MobilityInventory";
            var data = lineItems.ToDataSet(tableName);
            if (data.Tables?.Count == 0 || data.Tables.IndexOf(tableName) < 0)
            {
                return data;
            }
            if (!permissionManager.UserCanAccess(Session, ModuleEnum.RevCustomers))
            {
                data.Tables[tableName].Columns.Remove("Acct #");
            }

            if (!permissionManager.UserIsSuperAdmin(Session) && !permissionManager.IsParentAdmin)
            {
                data.Tables[tableName].Columns.Remove("Foundation Acct #");
                data.Tables[tableName].Columns.Remove("Billing Acct #");
                data.Tables[tableName].Columns.Remove("BAN Status");
                data.Tables[tableName].Columns.Remove("Data Group");
                data.Tables[tableName].Columns.Remove("Carrier Pool");
                data.Tables[tableName].Columns.Remove("Carrier Rate Plan");
                data.Tables[tableName].Columns.Remove("Carrier Rate Plan Cost");
                data.Tables[tableName].Columns.Remove("Carrier Data Allocaton MB");
                data.Tables[tableName].Columns["Customer Rate Plan"].ColumnName = "Rate Plan";
                data.Tables[tableName].Columns["Customer Rate Plan Cost"].ColumnName = "Rate Plan Cost";
                data.Tables[tableName].Columns["Customer Data Allocation MB"].ColumnName = "Data Allocation MB";
                data.Tables[tableName].Columns["Customer Pool"].ColumnName = "Pool";
            }

            if (permissionManager.OptimizationSettings.OptIntoCrossProviderCustomerOptimization)
            {
                data.Tables[tableName].Columns.Remove(CommonDisplayNames.DATA_USAGE_MB);
            }
            else
            {
                data.Tables[tableName].Columns.Remove(CommonDisplayNames.CARRIER_CYCLE_USAGE);
                data.Tables[tableName].Columns.Remove(CommonDisplayNames.CUSTOMER_CYCLE_USAGE);
            }

            if (!permissionManager.MultipleCostCenter)
            {

                data.Tables[tableName].Columns["Cost Center 1"].ColumnName = "Cost Center";
                data.Tables[tableName].Columns.Remove("Cost Center 2");
                data.Tables[tableName].Columns.Remove("Cost Center 3");
            }

            return data;
        }

        public ActionResult MobilityDeviceDetails(int deviceId)
        {
            ViewBag.PageTitle = "Mobility Device Details";

            if (permissionManager.UserCannotAccess(Session, ModuleEnum.Mobility))
                return RedirectToAction("Index", "Home");

            var deviceDetailsRepository = new MobilityInventoryRepository(altaWrxDb, permissionManager);

            var model = new MobilityDeviceModel
            {
                IsAdmin = permissionManager.UserIsSuperAdmin(Session) || permissionManager.IsParentAdmin,
                Device = deviceDetailsRepository.GetDevice(deviceId)
            };

            var device = model.Device;
            if (device != null)
            {
                model.CarrierPoolUsage = GetPoolUsageModel(device);
                model.CarrierDataGroupUsage = GetDataGroupUsageModel(device);

                if (device.IntegrationId == (int)IntegrationType.Telegence)
                {
                    var deviceUsageRepository = new TelegenceDeviceUsageRepository(altaWrxDb, device.ServiceProviderId);
                    model.UsageDetailRecords = deviceUsageRepository.GetUsageByDevice(device.MSISDN);
                }

                model.CustomerRatePoolUsage = device.CustomerRatePoolId != null
                    ? GetCustomerRatePoolUsage(device.CustomerRatePoolId.GetValueOrDefault())
                    : new CustomerRatePoolDetailsModel();

                model.MobilityFeatures = new List<MobilityFeatureModel>();
                if (!string.IsNullOrEmpty(model.Device.FeatureCodes))
                {
                    var mobilityFeatureRepo = new MobilityFeatureRepository(altaWrxDb);
                    model.MobilityFeatures = mobilityFeatureRepo.GetFeatureBySOCCode(permissionManager, Session, model.Device.FeatureCodes, device.ServiceProviderId).ToList();
                }
            }

            return PartialView("_MobilityDeviceDetails", model);
        }

        private CustomerRatePoolDetailsModel GetCustomerRatePoolUsage(int customerRatePoolId)
        {
            var customerRatePoolRepository = new CustomerRatePoolRepository(altaWrxDb, permissionManager);
            var mobilityDevices = customerRatePoolRepository.GetCustomerRatePoolLineDetails(PortalTypes.Mobility, customerRatePoolId, "");
            var totalDataAllocation = mobilityDevices.Sum(md => md.DataAllocationMB);
            var totalUsage = mobilityDevices.Sum(md => md.DataUsageMB);
            var percentUsage = totalDataAllocation > 0 ? (totalUsage / totalDataAllocation) * 100 : 0;
            var customerRatePoolDetails = new CustomerRatePoolDetailsModel
            {
                TotalDataAllocationMB = totalDataAllocation,
                TotalDataUsageMB = totalUsage,
                PercentDataUsage = percentUsage,
                Name = mobilityDevices.FirstOrDefault()?.CustomerRatePoolName
            };
            return customerRatePoolDetails;
        }

        public ActionResult GetDataGroupUsage(string dataGroupId, string billingAccountNo, string foundationAccountNo, int serviceProviderId)
        {
            var model = GetDataGroupUsageModel(dataGroupId, billingAccountNo, foundationAccountNo, serviceProviderId);
            return PartialView("_DataGroupUsage", model);
        }

        private MobilityAggregateUsageModel GetDataGroupUsageModel(string dataGroupId, string billingAccountNo, string foundationAccountNo,
            int serviceProviderId)
        {
            var repo = new MobilityDeviceUsageAggregateRepository(altaWrxDb, serviceProviderId);
            var model = new MobilityAggregateUsageModel
            {
                FoundationAccountNo = foundationAccountNo,
                BillingAccountNo = billingAccountNo,
                DataGroupId = dataGroupId,
                MobilityDeviceUsageAggregate = repo.GetDataGroupUsage(dataGroupId, billingAccountNo, foundationAccountNo),
                UsageForDisplay = 0.0M
            };

            if (model.MobilityDeviceUsageAggregate?.DataUsage != null)
            {
                model.UsageForDisplay = Utils.BytesToGB(model.MobilityDeviceUsageAggregate.DataUsage.Value);
            }

            return model;
        }

        private MobilityAggregateUsageModel GetDataGroupUsageModel(IMobilityInventoryResult device)
        {
            return GetDataGroupUsageModel(device.DataGroupId, device.BillingAccountNumber,
                device.FoundationAccountNumber, device.ServiceProviderId);
        }

        public ActionResult GetPoolUsage(string poolId, string billingAccountNo, string foundationAccountNo, int serviceProviderId)
        {
            var model = GetPoolUsageModel(poolId, billingAccountNo, foundationAccountNo, serviceProviderId);
            return PartialView("_CarrierPoolUsage", model);
        }

        private MobilityAggregateUsageModel GetPoolUsageModel(string poolId, string billingAccountNo, string foundationAccountNo, int serviceProviderId)
        {
            var repo = new MobilityDeviceUsageAggregateRepository(altaWrxDb, serviceProviderId);
            var model = new MobilityAggregateUsageModel
            {
                FoundationAccountNo = foundationAccountNo,
                BillingAccountNo = billingAccountNo,
                PoolId = poolId,
                MobilityDeviceUsageAggregate = repo.GetPoolUsage(poolId, billingAccountNo, foundationAccountNo),
                UsageForDisplay = 0.0M
            };

            if (model.MobilityDeviceUsageAggregate?.DataUsage != null)
            {
                model.UsageForDisplay = Utils.BytesToGB(model.MobilityDeviceUsageAggregate.DataUsage.Value);
            }

            return model;
        }

        private MobilityAggregateUsageModel GetPoolUsageModel(IMobilityInventoryResult device)
        {
            return GetPoolUsageModel(device.PoolId, device.BillingAccountNumber,
                device.FoundationAccountNumber, device.ServiceProviderId);
        }

        [HttpPost]
        public async Task<ActionResult> UpdateMobilityConfiguration(int deviceId, int serviceProviderId, string configurationTypeString, List<string> mobilityConfigurationIDsToAdd,
            List<string> mobilityConfigurationIDsToRemove, List<string> mobilityConfigurationsCurrent, string effectiveDate = "", string msisdn = null, int? optimizationGroup = null)
        {
            Enum.TryParse(configurationTypeString, out MobilityConfigurationType configurationType);
            var repository = new MobilityConfigurationChangeQueueRepository(altaWrxDb);
            TelegenceDeviceRepository TDR = new TelegenceDeviceRepository(altaWrxDb);
            try
            {
                var telegenceDeviceId = string.IsNullOrEmpty(msisdn) ? 0 : TDR.GetBySubscriberNumber(msisdn).id;

                var queue = CreateMobilityConfigurationChangeQueue(deviceId, serviceProviderId,
                    mobilityConfigurationIDsToAdd, mobilityConfigurationIDsToRemove, mobilityConfigurationsCurrent, configurationType, effectiveDate, telegenceDeviceId, optimizationGroup);

                var entityId = repository.SaveNew(queue);
                var result = await EnqueueMobilityConfigurationChangeAsync(entityId);

                if (string.IsNullOrEmpty(result))
                {
                    return new JsonResult { Data = "Success" };
                }
                return new JsonResult { Data = result };
            }
            catch (Exception e)
            {
                return new JsonResult { Data = e.Message };
            }
        }

        [HttpGet]
        public ActionResult SearchDevices(int? serviceProviderId = null, string filter = "")
        {
            if (!permissionManager.UserCanAccess(Session, ModuleEnum.Mobility))
            {
                return new HttpStatusCodeResult(HttpStatusCode.Forbidden);
            }

            var repository = new MobilityInventoryRepository(altaWrxDb, permissionManager);
            var devices = repository.SearchDevices(serviceProviderId, filter, 25);

            return Json(devices.Select(device => new { id = device.id, label = device.MSISDN, value = device.MSISDN }), JsonRequestBehavior.AllowGet);
        }


        public ActionResult BulkChanges(int? serviceProviderId = null, int? changeType = null, string filter = null, long qualificationId = 0, int page = 1, int pageSize = 25, string sort = null, string sortDir = null)
        {
            if (!permissionManager.UserCanAccessPortalTypeModule(Session, PORTAL_TYPE))
            {
                return new HttpStatusCodeResult(HttpStatusCode.Forbidden);
            }

            var serviceProviderRepository = new ServiceProviderRepository(altaWrxDb);
            var serviceProviders = serviceProviderRepository.GetAllByPortalType(PORTAL_TYPE);
            var integrationTypes = serviceProviders.Select(serviceProvider => (IntegrationType)serviceProvider.IntegrationId).ToArray();

            var changeTypeRepository = new DeviceChangeRequestTypeRepository(altaWrxDb);
            ICollection<DeviceChangeRequestType> changeTypes;
            if (serviceProviderId.HasValue)
            {
                var serviceProvider = serviceProviders.First(sp => sp.id == serviceProviderId.Value);
                changeTypes = changeTypeRepository.GetAllByIntegrationId(serviceProvider.IntegrationId, permissionManager.HighestRoleForUserInTenant());
            }
            else
            {
                changeTypes = changeTypeRepository.GetAllByIntegrationType(permissionManager.HighestRoleForUserInTenant(), integrationTypes);
            }

            var changeRepository = new DeviceBulkChangeRepository(altaWrxDb, permissionManager);
            var search = new BulkChangeSearch
            {
                ServiceProviderId = serviceProviderId,
                ChangeType = changeType,
                Filter = filter,
                PortalType = PORTAL_TYPE
            };
            var changes = changeRepository.GetBulkChanges(search, sort, sortDir, page, pageSize);

            var qualification = new QualificationViewModel();
            if (qualificationId > 0)
            {
                var qualificationRepository = new QualificationRepository(altaWrxDb);
                var qualificationEntity = qualificationRepository.GetByIdWithAddresses(qualificationId);
                qualificationEntity.QualificationAddresses = qualificationEntity.QualificationAddresses.Where(x => x.IsQualified && !x.IsActivatedByService).ToList();
                if (qualificationEntity != null)
                {
                    qualification.QualificationId = qualificationEntity.id;
                    qualification.ServiceLineId = qualificationEntity.ServiceLineId;
                    qualification.TenantId = qualificationEntity.TenantId;

                    if (qualificationEntity.QualificationAddresses.Count() > 0)
                    {
                        qualification.QualificationAddresses = qualificationEntity.QualificationAddresses.Select(x => new QualificationAddressViewModel
                        {
                            QualificationToken = x.QualificationToken,
                            StreetNumber = x.StreetNumber,
                            StreetName = x.StreetName,
                            City = x.City,
                            State = x.State,
                            Country = x.Country,
                            PostalCode = x.ZipCode,
                            SiteId = x.SiteId
                        }).ToList();
                    }

                    if (qualificationEntity.QualificationAddresses.Count() == 1)
                    {
                        qualification.DefaultQualificationAddressToken = qualificationEntity.QualificationAddresses.First().QualificationToken;
                    }
                }
            }

            TimeZoneHelper.ApplySystemTimezoneForBulkChangeDateTimeColumn(changes, permissionManager.AltaworxCentralConnectionString, out string simpleTimeZoneInfoName);
            var totalCount = changeRepository.GetBulkChangeCount(search);
            var model = new BulkChangeListViewModel
            {
                PageSize = pageSize,
                ServiceProviderId = serviceProviderId,
                ChangeType = changeType,
                Filter = filter,
                ServiceProviders = serviceProviders,
                ChangeTypes = changeTypes,
                Changes = PagedList.ToPagedList(changes, totalCount),
                States = ListHelper.States(),
                IsBulkChangeRunning = changeRepository.CheckBulkChangeIsRunning((int)PortalTypes.Mobility),
                TimeZoneInfoName = simpleTimeZoneInfoName,
                Qualification = qualification
            };

            return View("BulkChanges", model);
        }

        [HttpGet]
        public ActionResult BulkChange(long id, int page = 1, int pageSize = 25, string sort = null, string sortDir = null)
        {
            if (!permissionManager.UserCanAccessPortalTypeModule(Session, PORTAL_TYPE))
            {
                return new HttpStatusCodeResult(HttpStatusCode.Forbidden);
            }

            var changeRepository = new DeviceBulkChangeRepository(altaWrxDb, permissionManager);
            var bulkChange = changeRepository.GetBulkChange(id);

            if (bulkChange == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.NotFound);
            }

            var serviceProviderRepository = new ServiceProviderRepository(altaWrxDb);
            var serviceProviders = serviceProviderRepository.GetAllByPortalType(PORTAL_TYPE);
            var integrationTypes = serviceProviders.Select(serviceProvider => (IntegrationType)serviceProvider.IntegrationId).ToArray();

            var changeTypeRepository = new DeviceChangeRequestTypeRepository(altaWrxDb);
            ICollection<DeviceChangeRequestType> changeTypes;
            if (bulkChange.ServiceProviderId > 0)
            {
                var serviceProvider = serviceProviders.First(sp => sp.id == bulkChange.ServiceProviderId);
                changeTypes = changeTypeRepository.GetAllByIntegrationId(serviceProvider.IntegrationId, permissionManager.HighestRoleForUserInTenant());
            }
            else
            {
                changeTypes = changeTypeRepository.GetAllByIntegrationType(permissionManager.HighestRoleForUserInTenant(), integrationTypes);
            }

            var mobilityDeviceChangeRepository = new MobilityDeviceChangeRepository(altaWrxDb, permissionManager);
            var changes = mobilityDeviceChangeRepository.GetChanges(id, sort, sortDir, page, pageSize);
            var changeCount = mobilityDeviceChangeRepository.GetCountForBulkChange(id);

            var qualificationRepository = new QualificationRepository(altaWrxDb);
            var qualification = qualificationRepository.GetByDeviceBulkChangeId(id);

            var model = new BulkChangeDetailsViewModel
            {
                bulkChangeId = id,
                Count = bulkChange.ChangeCount.GetValueOrDefault(),
                ServiceProviderId = bulkChange.ServiceProviderId,
                ChangeType = bulkChange.ChangeRequestTypeId,
                ProcessedCount = bulkChange.ProcessedCount.GetValueOrDefault(),
                ErrorCount = bulkChange.ErrorCount.GetValueOrDefault(),
                Status = bulkChange.Status,
                ProcessedDate = bulkChange.ProcessedDate,
                ServiceProviders = serviceProviders,
                ChangeTypes = changeTypes,
                Details = PagedList.ToPagedList(changes, changeCount),
                States = ListHelper.States(),
                Qualification = qualification,
            };
            model.SubcriberNumbers = string.Join(@"\n", model.Details.TheList.Select(item => item.SubscriberNumber).Distinct());
            return View("BulkChange", model);
        }

        [HttpGet]
        public ActionResult ChangeICCIDorIMEI(int serviceProviderId)
        {
            if (!permissionManager.UserCanAccessPortalTypeModule(Session, PORTAL_TYPE))
            {
                return new HttpStatusCodeResult(HttpStatusCode.Forbidden);
            }
            BulkchangeUpdateICCIDorIMEI model = new BulkchangeUpdateICCIDorIMEI()
            {
                ServiceProviderId = serviceProviderId,
                PortalType = PortalTypes.Mobility
            };
            return PartialView("_BulkChangeUpdateICCIDorIMEI", model);
        }
        [HttpPost]
        public async Task<ActionResult> PostChangeICCIDorIMEI(BulkchangeUpdateICCIDorIMEI model)
        {
            if (!permissionManager.UserCanAccessPortalTypeModule(Session, PORTAL_TYPE))
            {
                return new HttpStatusCodeResult(HttpStatusCode.Forbidden);
            }
            try
            {
                model.Devices = model.Devices.Distinct().ToList();
                var changeRepository = new DeviceBulkChangeRepository(altaWrxDb, permissionManager);
                var changeType = DeviceChangeType.ChangeICCIDorIMEI;
                var bulkChange = new DeviceBulkChange
                {
                    ChangeRequestTypeId = (int)changeType,
                    ServiceProviderId = model.ServiceProviderId,
                    TenantId = permissionManager.PermissionFilter.LoggedInTenantId,
                    Status = BulkChangeStatus.NEW,
                    CreatedDate = DateTime.UtcNow,
                    CreatedBy = SessionHelper.GetAuditByName(Session),
                    IsActive = true,
                    IsDeleted = false,
                    Mobility_DeviceChange =
                        BuildUpdateICCIDorIMEI(altaWrxDb, Session, permissionManager, model).ToList()
                };
                var bulkChangeId = changeRepository.CreateBulkChange(bulkChange);
                await ProcessBulkChange(bulkChange.id);
                return new JsonResult { Data = new { Success = true, ChangeId = bulkChangeId } };
            }
            catch (Exception e)
            {
                return new JsonResult { Data = new { Success = false, e.Message } };
            }
            //return Json(new { status = "OK" }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult ValidateBulkChange(BulkChangeCreateModel bulkChangeCreateModel)
        {
            if (!permissionManager.UserCanCreate(Session, ModuleEnum.Mobility))
            {
                return new HttpStatusCodeResult(HttpStatusCode.Forbidden);
            }

            if (!ModelState.IsValid)
            {
                return Json(new
                {
                    isValid = false,
                    validationMessage = "Errors Validating Model",
                    errors = ModelState.SelectMany(modelState => modelState.Value.Errors).Select(error => error.ErrorMessage).ToArray()
                });
            }

            var changeRepository = new DeviceBulkChangeRepository(altaWrxDb, permissionManager);
            var serviceProviderId = bulkChangeCreateModel.ServiceProviderId.GetValueOrDefault();
            var changeTypeId = bulkChangeCreateModel.ChangeType.GetValueOrDefault();
            var changeType = (DeviceChangeType)changeTypeId;
            var changes = BuildChangeDetails(bulkChangeCreateModel, serviceProviderId, changeType).ToList();
            var changesWithErrors = changes.Where(x => x.HasErrors && x.StatusDetails.StartsWith("Active Rev Services")).ToList();
            if (changesWithErrors.Count > 0)
            {
                return Json(new
                {
                    isValid = false,
                    validationMessage = "One or more devices have active Rev Services",
                    errors = changesWithErrors.Select(x => x.SubscriberNumber).ToArray()
                });
            }
            else
            {
                return Json(new { isValid = true });
            }
        }

        [HttpPost]
        public async Task<ActionResult> BulkChange(BulkChangeCreateModel bulkChangeCreateModel)
        {
            //TODO update model & validate the new fields
            if (!permissionManager.UserCanCreate(Session, ModuleEnum.Mobility))
            {
                return new HttpStatusCodeResult(HttpStatusCode.Forbidden);
            }

            if (!ModelState.IsValid)
            {
                return Json(new
                {
                    errors = ModelState.SelectMany(modelState => modelState.Value.Errors).Select(error => error.ErrorMessage).ToArray()
                });
            }
            var auditUser = SessionHelper.GetAuditByName(Session);
            var changeRepository = new DeviceBulkChangeRepository(altaWrxDb, permissionManager);
            var serviceProviderId = bulkChangeCreateModel.ServiceProviderId.GetValueOrDefault();
            var changeTypeId = bulkChangeCreateModel.ChangeType.GetValueOrDefault();
            var changeType = (DeviceChangeType)changeTypeId;
            var serviceProviderRepository = new ServiceProviderRepository(altaWrxDb);
            var serviceProvider = serviceProviderRepository.GetById(serviceProviderId);
            var integrationType = serviceProvider.IntegrationId;
            var billingAccountNumber = bulkChangeCreateModel?.StatusUpdate?.TelegenceStatusUpdate?.BillingAccountNumber;

            if (changeType.Equals(DeviceChangeType.ActivateNewService) && integrationType == (int)IntegrationEnum.Telegence)
            {
                var isTelegenceBanValid = await IsTelegenceBillingAccountNumberValid(serviceProviderId, billingAccountNumber);
                if (!isTelegenceBanValid)
                {
                    return Json(new
                    {
                        errors = new List<string>() { string.Format(CommonStrings.BillingAccountNumberCouldNotBeFound, billingAccountNumber) }
                    });
                }
            }

            var changes = BuildChangeDetails(bulkChangeCreateModel, serviceProviderId, changeType).ToList();

            var bulkChange = new DeviceBulkChange
            {
                ChangeRequestTypeId = changeTypeId,
                ServiceProviderId = serviceProviderId,
                TenantId = permissionManager.PermissionFilter.LoggedInTenantId,
                SiteId = GetSiteIdForBulkChange(changes),
                Status = changes.Any(change => !change.IsProcessed) ? BulkChangeStatus.NEW : BulkChangeStatus.PROCESSED,
                CreatedDate = DateTime.UtcNow,
                CreatedBy = auditUser,
                IsActive = true,
                IsDeleted = false,
                Mobility_DeviceChange = changes
            };

            var bulkchangeAssignCustomer = new DeviceBulkChange();
            var iccidsChanges = changes.Where(x => !x.HasErrors).Select(x => x.ICCID).ToList();

            try
            {
                var bulkchangeId = changeRepository.CreateBulkChange(bulkChange);
                if (bulkchangeId > 0 && DeviceChangeType.ActivateNewService == changeType)
                {
                    var qualificationRepository = new QualificationRepository(altaWrxDb);
                    qualificationRepository.MarkQualificationAddressesAsActivated(bulkchangeId, auditUser);
                }
                if (DeviceChangeType.ActivateNewService == changeType && !string.IsNullOrWhiteSpace(bulkChangeCreateModel.StatusUpdate.TelegenceStatusUpdate.RevAccountNumber) && iccidsChanges.Any())
                {
                    // get msisdn List
                    var deviceRepo = new MobilityDeviceRepository(altaWrxDb);
                    BulkChangeAssociateCustomerModel model = new BulkChangeAssociateCustomerModel()
                    {
                        ServiceProviderId = bulkChangeCreateModel.ServiceProviderId.Value,
                        RevCustomerId = bulkChangeCreateModel.StatusUpdate.TelegenceStatusUpdate.RevAccountNumber,
                        Devices = iccidsChanges.ToArray(),
                        AddCarrierRatePlan = false,
                        AddCustomerRatePlan = false,
                        Prorate = false,
                        CreateRevService = false
                    };

                    var mobilityChanges = BuildAssociateCustomerChangeAfterActivation(altaWrxDb, Session, permissionManager, model, true);
                    bulkchangeAssignCustomer = BuildBulkChange(serviceProviderId, mobilityChanges.ToList(), DeviceChangeType.CustomerAssignment);
                    bulkchangeAssignCustomer.Status = BulkChangeStatus.PROCESSING;
                    changeRepository.CreateBulkChange(bulkchangeAssignCustomer);
                }

                if (bulkchangeAssignCustomer.id > 0 && bulkChangeCreateModel.ProcessChanges.GetValueOrDefault())
                {
                    return await ProcessBulkChange(bulkChange.id, bulkchangeAssignCustomer.id);
                }
                else if (bulkChangeCreateModel.ProcessChanges.GetValueOrDefault())
                {
                    return await ProcessBulkChange(bulkChange.id);
                }

                return Json(new { redirectUrl = Url.Action("BulkChange", "Mobility", new { bulkChange.id }) }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { errors = new[] { $"An error occurred: {ex.Message}" } }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public async Task<ActionResult> BulkChangeCarrierRatePlan(FormCollection collection)
        {
            var uploadedFile = Request.Files[0];
            var serviceProviderId = Convert.ToInt32(collection["DropZoneServiceProviderId"]);
            var integrationId = Convert.ToInt32(collection["DropZoneIntegrationId"]);
            if (uploadedFile?.FileName == null)
            {
                return ErrorMessage("Empty file name. Must select a valid file to process.");
            }

            if (Path.GetExtension(uploadedFile.FileName).ToUpper() != ".CSV")
            {
                return ErrorMessage($"Invalid File: {uploadedFile.FileName}.  The file must be in .CSV format.");
            }

            if (uploadedFile.FileName.Length > 255)
            {
                return ErrorMessage(
                    $"Invalid File: {uploadedFile.FileName}.  The filename must be less than 255 characters long.");
            }

            try
            {
                var serviceProviderRepository = new ServiceProviderRepository(altaWrxDb);
                var serviceProviders = serviceProviderRepository.GetAll();

                if (serviceProviders.Count(sp => sp.id == Convert.ToInt32(serviceProviderId)) == 0)
                    return ErrorMessage($"Invalid Service Provider Id: {serviceProviderId}");

                using (var streamReader = new StreamReader(uploadedFile.InputStream))
                {
                    var csvConfiguration = new CsvConfiguration(CultureInfo.InvariantCulture, prepareHeaderForMatch: Helpers.CsvHelper.PrepareHeaderForMatch);
                    using (var csv = new CsvReader(streamReader, csvConfiguration))
                    {
                        var lines = csv.GetRecords<CarrierRatePlanChangeCsvRow>().ToList();
                        if (!lines.Any())
                        {
                            return ErrorMessage($"No rows found in '{uploadedFile.FileName}'.");
                        }

                        var mobilityDeviceRepository = new MobilityDeviceRepository(altaWrxDb);
                        var mobilityDevices = mobilityDeviceRepository.GetAllByServiceProviderId(serviceProviderId);

                        var validationResult = ValidateCarrierRatePlanChange(serviceProviderId, lines, mobilityDevices);
                        if (!string.IsNullOrEmpty(validationResult))
                            return ErrorMessage($"Invalid Carrier Rate Plan Upload: {validationResult}");

                        var processResult = await ProcessCarrierRatePlanChange(serviceProviderId, integrationId, lines, mobilityDevices);
                        if (!string.IsNullOrEmpty(processResult))
                            return ErrorMessage($"{processResult}");
                    }
                }

            }
            catch (Exception ex)
            {
                return ErrorMessage($"Error: {ex.Message}");
            }
            SessionHelper.SetAlert(Session, "Successfully uploaded carrier rate plan.");
            SessionHelper.SetAlertType(Session, "success");
            return Content("OK");
        }
        [HttpPost]
        public ActionResult GetRatePlanCodeByIMEI(string[] devices, int? qualificationId = null)
        {
            var changesWithErrors = new List<Mobility_DeviceChange>();
            var createdBy = SessionHelper.GetAuditByName(Session);
            var deviceList = devices.Where(d => !string.IsNullOrWhiteSpace(d)).ToList();
            if (deviceList.Count > 0)
            {
                var deviceItems = new List<Device>();
                deviceItems = GetDeviceItems(deviceList, changesWithErrors, createdBy);
                if (deviceItems.Count > 0 && changesWithErrors.Count == 0)
                {
                    var listIMEI = deviceItems.Select(x => x.IMEI).ToArray();
                    var deviceTypeByIMEIs = GetDeviceTypeByIMEI(string.Join(",", listIMEI));
                    if (deviceTypeByIMEIs != null && deviceTypeByIMEIs.Count > 0)
                    {
                        var siteGroupCustomerRatePlans = GetJasperCustomerRatePlansBySideIdAndSideGroupId();
                        var isMultipleDevice = deviceList.Count > 1;
                        var compatibleRatePlansForSingleDevice = new List<CompatibleRatePlanSelectListItem>();
                        var compatibleRatePlanList = GetCompatibleRatePlanList(deviceTypeByIMEIs, deviceItems, siteGroupCustomerRatePlans, qualificationId);
                        if (!isMultipleDevice)
                        {
                            compatibleRatePlansForSingleDevice = compatibleRatePlanList.FirstOrDefault().RatePlans;
                        }
                        var bulkChangeActivationInformation = GetActivationInformation(compatibleRatePlansForSingleDevice, changesWithErrors, createdBy);
                        return Json(new
                        {
                            isValid = changesWithErrors.Count == 0,
                            isMultipleDevice = isMultipleDevice,
                            CompatibleRatePlans = compatibleRatePlanList,
                            ActivateInformation = bulkChangeActivationInformation,
                            errors = changesWithErrors.Where(x => x.HasErrors).Select(x => x.StatusDetails).ToArray()
                        });
                    }
                }
            }
            else
            {
                changesWithErrors.Add(CreateDeviceChangeError(null, Resources.CommonStrings.ICCIDCannotBeEmpty, createdBy));
            }
            if (changesWithErrors.Count > 0)
            {
                return Json(new
                {
                    isValid = false,
                    validationMessage = "ERROR",
                    errors = changesWithErrors.Where(x => x.HasErrors).Select(x => x.StatusDetails).ToArray()
                });
            }
            return Json(new { isValid = true });
        }
        private BulkChangeActivationInformation GetActivationInformation(List<CompatibleRatePlanSelectListItem> ratePlans, List<Mobility_DeviceChange> changesWithErrors, string createdBy)
        {
            var bulkChangeActivationInformation = new BulkChangeActivationInformation();
            var userRoleSite = GetUserRoleSite();
            var siteId = userRoleSite.SiteId;
            var siteGroupId = userRoleSite.SiteGroupId;
            var siteGroup = GetSiteGroupByCurrentUser(siteId, siteGroupId);
            if (siteGroup != null && !string.IsNullOrWhiteSpace(siteGroup.BillingAccountNumber))
            {
                var assignCustomer = ListHelper.RevCustomerList(permissionManager, permissionManager.Tenant.id, permissionManager.PermissionFilter.GetRevAccountFilter(), string.Empty);
                bulkChangeActivationInformation.BillingAccountNumber = siteGroup.BillingAccountNumber;
                bulkChangeActivationInformation.AssignCustomer = assignCustomer;
            }
            else
            {
                changesWithErrors.Add(CreateDeviceChangeError(null, string.Format(Resources.CommonStrings.WarningContactAdmin, "Customer Group"), createdBy));
            }
            bulkChangeActivationInformation.RatePlans = ratePlans;
            return bulkChangeActivationInformation;
        }

        private List<JasperCustomerRatePlan> GetJasperCustomerRatePlansBySideIdAndSideGroupId()
        {
            var userRoleSite = GetUserRoleSite();
            var siteId = userRoleSite.SiteId;
            var siteGroupId = userRoleSite.SiteGroupId;
            var siteGroup = GetSiteGroupByCurrentUser(siteId, siteGroupId);
            if (siteGroup != null && siteGroup.SiteGroup_CustomerRatePlan.Count > 0)
            {
                return siteGroup.SiteGroup_CustomerRatePlan.Select(x => x.JasperCustomerRatePlan).ToList();
            }
            return new List<JasperCustomerRatePlan>();
        }

        private User_Role_Site GetUserRoleSite()
        {
            User activeUser = SessionHelper.User(Session);
            var userRoleSite = activeUser.User_Role_Site.FirstOrDefault(x => x.TenantId == permissionManager.Tenant.id && x.IsActive && !x.IsDeleted);
            if (permissionManager.IsChildAdmin)
            {
                userRoleSite = activeUser.User_Role_Site.FirstOrDefault(x => x.TenantId == permissionManager.Tenant.ParentTenantId && x.IsActive && !x.IsDeleted);
            }
            return userRoleSite;
        }

        private List<SelectListItem> GetAssignCustomers(int? siteGroupId, int? siteId)
        {
            var assignCustomer = new List<SelectListItem>();
            if (siteGroupId == null && siteId.HasValue)
            {
                var site = altaWrxDb.Sites.FirstOrDefault(x => x.id == siteId.Value);
                assignCustomer.Add(new SelectListItem
                {
                    Value = site.id.ToString(),
                    Text = site.DisplayName
                });
            }
            else
            {
                var siteIds = altaWrxDb.SiteGroup_Site.Where(x => x.SiteGroupId == siteGroupId.Value).Select(x => x.SiteId).ToList();
                var sites = altaWrxDb.Sites.Where(x => siteIds.Contains(x.id)).ToList();
                if (sites.Count > 0)
                {
                    assignCustomer = sites.Select(x => new SelectListItem
                    {
                        Value = x.id.ToString(),
                        Text = x.DisplayName
                    }).ToList();
                }
            }
            return assignCustomer;
        }
        private List<Device> GetDeviceItems(List<string> deviceList, List<Mobility_DeviceChange> changesWithErrors, string createdBy)
        {
            var deviceItems = new List<Device>();
            foreach (var device in deviceList)
            {
                var ids = Regex.Split(device.Trim(), @"[^\d]+").Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
                if (ids.Length != 2)
                {
                    changesWithErrors.Add(CreateDeviceChangeError(null, $"Invalid input line: {device}", createdBy));
                    continue;
                }
                var iccid = ids[0];
                var imei = ids[1];

                deviceItems.Add(new Device
                {
                    ICCID = iccid,
                    IMEI = imei,
                });
            }
            return deviceItems;
        }
        private List<DeviceBulkChangeCompatibleRatePlan> GetCompatibleRatePlanList(List<IMEIDeviceType> deviceTypeByIMEIs, List<Device> deviceItems, List<JasperCustomerRatePlan> siteGroupCustomerRatePlans, int? qualificationId = null)
        {
            var compatibleRatePlanList = new List<DeviceBulkChangeCompatibleRatePlan>();
            foreach (var item in deviceTypeByIMEIs)
            {
                var ratePlanSelected = new List<CompatibleRatePlanSelectListItem>();
                var ratePlansQuery = altaWrxDb.JasperCarrierRatePlans.Where(x => x.IMEIType.Contains(item.DeviceType));
                if (qualificationId != null && qualificationId > 0)
                {
                    ratePlansQuery = ratePlansQuery
                        .Where(rp => rp.Family != null)
                        .Where(rp => CarrierRatePlanEnum.FIXED_WIRELESS_FAMILIES.Contains(rp.Family.Trim()));
                }

                var ratePlans = ratePlansQuery.ToList();
                if (ratePlans.Count > 0)
                {
                    var ratePlanIds = ratePlans.Select(x => x.id);
                    var customerCarrierRatePlans = altaWrxDb.CustomerRatePlan_JasperCarrierRatePlan.Where(x => ratePlanIds.Contains(x.JasperCarrierRatePlanId)).ToList();
                    if (customerCarrierRatePlans.Count > 0)
                    {
                        ratePlanSelected = MapCustomerRatePlanCarrierRatePlan(siteGroupCustomerRatePlans, customerCarrierRatePlans, ratePlans);
                    }
                }
                var deviceItem = deviceItems.FirstOrDefault(x => x.IMEI == item.IMEI);
                var compatibleRatePlans = new DeviceBulkChangeCompatibleRatePlan
                {
                    Device = $"{deviceItem.ICCID},{deviceItem.IMEI}",
                    RatePlans = ratePlanSelected
                };
                compatibleRatePlanList.Add(compatibleRatePlans);
            }
            return compatibleRatePlanList;
        }
        private List<CompatibleRatePlanSelectListItem> MapCustomerRatePlanCarrierRatePlan(List<JasperCustomerRatePlan> customerRatePlans, List<CustomerRatePlan_JasperCarrierRatePlan> customerCarrierRatePlans, List<JasperCarrierRatePlan> ratePlans)
        {
            var ratePlanSelected = new List<CompatibleRatePlanSelectListItem>();
            foreach (var customerRatePlan in customerRatePlans)
            {
                var ratePlanId = customerCarrierRatePlans.FirstOrDefault(x => x.CustomerRatePlanId == customerRatePlan.id)?.JasperCarrierRatePlanId;
                var socCode = ratePlans.FirstOrDefault(x => x.id == ratePlanId)?.RatePlanCode;
                if (ratePlanId.HasValue && !string.IsNullOrWhiteSpace(socCode))
                {
                    ratePlanSelected.Add(new CompatibleRatePlanSelectListItem
                    {
                        Value = socCode,
                        Text = customerRatePlan.RatePlanName,
                        CustomerRatePlanId = customerRatePlan.id
                    });
                }
            }
            return ratePlanSelected;
        }
        private List<IMEIDeviceType> GetDeviceTypeByIMEI(string imeis)
        {
            try
            {
                var deviceTypeByImeis = new List<IMEIDeviceType>();
                using (var conn = new SqlConnection(permissionManager.AltaworxCentralConnectionStringWithoutEF))
                {
                    using (var cmd = new SqlCommand("usp_DeviceBulkChange_Get_DeviceType_IMEI", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@IMEIS", imeis);
                        cmd.CommandTimeout = SQLConstant.TimeoutSeconds;
                        conn.Open();
                        var dataReader = cmd.ExecuteReader();
                        if (dataReader.HasRows)
                        {
                            while (dataReader.Read())
                            {
                                deviceTypeByImeis.Add(new IMEIDeviceType
                                {
                                    IMEI = dataReader["IMEI"].ToString(),
                                    DeviceType = dataReader["DeviceType"].ToString()
                                });
                            }
                        }
                    }
                }
                return deviceTypeByImeis;
            }
            catch (Exception ex)
            {
                Log.Error($"Error Executing Stored Procedure usp_DeviceBulkChange_Get_DeviceType_IMEI: {ex.Message} {ex.StackTrace}", ex);
                return null;
            }
        }
        private string ValidateCarrierRatePlanChange(int serviceProviderId, IList<CarrierRatePlanChangeCsvRow> lines, IList<Models.Repositories.MobilityDevice> mobilityDevices)
        {
            var sb = new StringBuilder();

            var jasperCarrierRatePlanRepository = new JasperCarrierRatePlanRepository(altaWrxDb);
            var ratePlans = jasperCarrierRatePlanRepository.GetAllByServiceProvider(serviceProviderId);

            var billingPeriodRepository = new BillingPeriodRepository(altaWrxDb);
            var currentBillingPeriodStartDate = billingPeriodRepository.GetCurrentBillingPeriodByServiceProviderId(serviceProviderId);

            for (var i = 0; i < lines.Count; i++)
            {
                var line = lines[i];
                if (!mobilityDevices.Any(md => md.MSISDN == line.PhoneNumber))
                {
                    sb.AppendLine($"Invalid Phone number {line.PhoneNumber} on Row {i + 1}");
                    continue;
                }
                if (!ratePlans.Any(md => md.RatePlanCode == line.RatePlanName))
                {
                    sb.AppendLine($"Invalid Rate Plan {line.RatePlanName} on Row {i + 1}");
                    continue;
                }

                if (line.EffectiveDate != null && !ValidateCarrierRatePlanEffectiveDate(line.EffectiveDate.GetValueOrDefault(),
                        currentBillingPeriodStartDate.GetValueOrDefault()))
                {
                    sb.AppendLine(
                        $"Invalid Effective Date {line.EffectiveDate} on Row {i + 1}. Date cannot be after 30 days and not before the first day of the current billing cycle");
                }
            }

            return sb.ToString();
        }

        //bulkchange activation information 
        [HttpPost]
        public async Task<ActionResult> BulkChangeActivationInformation(FormCollection collection)
        {
            User activeUser = SessionHelper.User(Session);
            var userRoleSite = activeUser.User_Role_Site.FirstOrDefault();
            var uploadedFile = Request.Files[0];
            var serviceProviderId = Convert.ToInt32(collection["DropZoneServiceProviderId1"]);
            var integrationId = Convert.ToInt32(collection["DropZoneIntegrationId1"]);
            if (uploadedFile?.FileName == null)
            {
                return Json(new { error = Resources.CommonStrings.FileEmpty });
            }

            if (Path.GetExtension(uploadedFile.FileName).ToUpper() != ".XLSX")
            {
                return Json(new { error = string.Format(Resources.CommonStrings.NameFileFormat, uploadedFile.FileName) });
            }

            if (uploadedFile.FileName.Length > MAXIMUM_CHARACTER_FILE_NAME)
            {
                return Json(new
                {
                    error = string.Format(Resources.CommonStrings.NameFileLong, uploadedFile.FileName)
                });
            }
            try
            {
                var serviceProviderRepository = new ServiceProviderRepository(altaWrxDb);
                var serviceProviders = serviceProviderRepository.GetAll();
                var changeRepository = new DeviceBulkChangeRepository(altaWrxDb, permissionManager);

                if (serviceProviders.Count(sp => sp.id == Convert.ToInt32(serviceProviderId)) == 0)
                    return Json(new { error = string.Format(Resources.CommonStrings.ServiceProviderIdInvalId, serviceProviderId) });
                using (ExcelPackage package = new ExcelPackage(uploadedFile.InputStream))
                {
                    ExcelWorkbook workbook = package.Workbook;
                    if (workbook != null)
                    {
                        ExcelWorksheet worksheet = workbook.Worksheets.FirstOrDefault();

                        if (worksheet != null)
                        {
                            // delete row empty
                            worksheet = ExcelUtilities.TrimEmptyRows(worksheet);
                            var activationCsvLines = new List<ActivationInformationChangeCsvRow>();
                            if (!permissionManager.IsAgent)
                            {
                                activationCsvLines = worksheet.ReadExcelToList<ActivationInformationChangeCsvRow>();
                            }
                            else
                            {
                                var billingAccountNumber = string.Empty;
                                var siteGroup = GetSiteGroupByCurrentUser(userRoleSite.SiteId, userRoleSite.SiteGroupId);
                                if (siteGroup != null)
                                {
                                    billingAccountNumber = siteGroup.BillingAccountNumber;
                                }
                                else
                                {
                                    return Json(new { error = string.Format(string.Format(Resources.CommonStrings.WarningContactAdmin, Resources.CommonStrings.CustomerGroup), uploadedFile.FileName) });
                                }

                                var activationCsvAgentLines = worksheet.ReadExcelToList<ActivationInformationChangeCsvRowAgent>();
                                activationCsvLines = activationCsvAgentLines.Select(item => item.ToActivationInformationChangeCsvRow(billingAccountNumber)).ToList();
                            }

                            // sheet two optional
                            if (!activationCsvLines.Any())
                            {
                                return Json(new { error = string.Format(Resources.CommonStrings.FileNotFound, uploadedFile.FileName) });
                            }

                            // validate excel field
                            var validationResult = new List<ValidateExcelUpload>();
                            validationResult = ValidateActivationInformationChange(serviceProviderId, activationCsvLines);

                            long BulkChangeAssignCustomerId = 0;
                            if (!permissionManager.IsAgent)
                            {
                                // If the user is the role Admin or SuperAdmin, the Excel file must 2 sheets. And Agent only needs 1 sheet
                                if (workbook.Worksheets.Count < NUMBER_OF_EXCEL_SHEETS)
                                {
                                    return Json(new { errors = new[] { Resources.CommonStrings.BulkChangeValidateExcelFile } }, JsonRequestBehavior.AllowGet);
                                }
                                ExcelWorksheet workSheetTwo = workbook.Worksheets[POSITION_OF_WORKSHEET_TWO];
                                workSheetTwo = ExcelUtilities.TrimEmptyRows(workSheetTwo);
                                var linesSheetTwo = workSheetTwo.ReadExcelToList<AssignCustomerCsvRow>();
                                if (linesSheetTwo != null && linesSheetTwo.Count > 0)
                                {
                                    var validationSheetTwoResult = new List<ValidateExcelUpload>();
                                    // validate sheet two
                                    validationSheetTwoResult = ValidateAssignCustomer(serviceProviderId, linesSheetTwo);
                                    if (validationSheetTwoResult.Count > 0)
                                    {
                                        return Json(new { errorSheetTwo = validationSheetTwoResult });
                                    }
                                    else
                                    {
                                        // build bulk change Assign customer
                                        var bulkchangeAssignCustomers = MappingRowToBulkchangeAssignCustomerModel(linesSheetTwo, serviceProviderId);
                                        if (bulkchangeAssignCustomers.Count > 0)
                                        {
                                            var mobilityChanges = new List<Mobility_DeviceChange>();
                                            foreach (var item in bulkchangeAssignCustomers)
                                            {
                                                mobilityChanges.AddRange(BuildAssociateCustomerByUploadFile(altaWrxDb, Session, permissionManager, item, (bool)item.IsUseCarrierActivation));
                                            }
                                            var bulkChangeAssignCustomer = BuildBulkChange(serviceProviderId, mobilityChanges, DeviceChangeType.CustomerAssignment);
                                            try
                                            {
                                                bulkChangeAssignCustomer.Status = BulkChangeStatus.PROCESSED;
                                                changeRepository.CreateBulkChange(bulkChangeAssignCustomer);
                                                if (bulkChangeAssignCustomer.id > 0)
                                                {
                                                    BulkChangeAssignCustomerId = bulkChangeAssignCustomer.id;
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                Log.Error(ex.Message, ex);
                                            }
                                        }
                                    }
                                }
                            }

                            if (validationResult.Count > 0)
                                return Json(new { error = validationResult });

                            // get devices change
                            var changes = BuildNewServiceActivationTelegenceByExcelFile(activationCsvLines).ToList();

                            // build bulk change
                            var bulkChange = BuildBulkChange(serviceProviderId, changes, DeviceChangeType.ActivateNewService);

                            try
                            {
                                changeRepository.CreateBulkChange(bulkChange);
                                if (bulkChange.id > 0)
                                {
                                    var qualificationRepo = new QualificationRepository(altaWrxDb);
                                    var processedBy = SessionHelper.GetAuditByName(Session);
                                    qualificationRepo.MarkQualificationAddressesAsActivated(bulkChange.id, processedBy);

                                    return await ProcessBulkChange(bulkChange.id, BulkChangeAssignCustomerId);
                                }

                                return Json(new { redirectUrl = Url.Action("BulkChange", "Mobility", new { bulkChange.id }), status = "OK" }, JsonRequestBehavior.AllowGet);
                            }
                            catch (Exception ex)
                            {
                                return Json(new { errors = new[] { $"An error occurred: {ex.Message}" } }, JsonRequestBehavior.AllowGet);
                            }
                        }
                    }
                }
                return Content("OK");
            }
            catch (Exception ex)
            {
                var message = $"Error bulk change activation information uploa by file excel: {ex.Message}";
                Log.Error(message, ex);
                return Json(new { error = $"Error: {ex.Message}" });
            }
        }

        private SiteGroup GetSiteGroupByCurrentUser(int? siteId, int? siteGroupId)
        {
            var siteGroupRepository = new SiteGroupRepository(altaWrxDb, permissionManager);
            if (siteId.HasValue)
            {
                var siteGroupIds = siteGroupRepository.GetSiteGroupIdsBySiteId(siteId.Value);
                return siteGroupRepository.GetSiteGroupHasBillingAccountNumberByIds(siteGroupIds);
            }
            if (siteGroupId.HasValue)
            {
                return siteGroupRepository.GetSiteGroup(siteGroupId.Value);
            }
            return null;
        }

        private List<BulkChangeAssociateCustomerModel> MappingRowToBulkchangeAssignCustomerModel(List<AssignCustomerCsvRow> linesSheet2, int serviceProviderId)
        {
            var result = new List<BulkChangeAssociateCustomerModel>();

            foreach (var line in linesSheet2)
            {
                if (result.Any(x => x.RevCustomerId == line.RevioCustomer))
                {
                    var itemUpdate = result.FirstOrDefault(x => x.RevCustomerId == line.RevioCustomer);
                    if (itemUpdate != null)
                    {
                        itemUpdate.RateList.Add(decimal.Parse(line.Rate));
                        itemUpdate.RevProductIdList.Add(int.Parse(line.RevioProduct));
                        var Iccids = itemUpdate.Devices.ToList();
                        Iccids.Add(line.ICCID);
                        itemUpdate.Devices = Iccids.ToArray();
                    }
                }
                else
                {
                    var effectiveDate = !string.IsNullOrWhiteSpace(line.EffectiveDate) ? DateTime.Parse(line.EffectiveDate) : DateTime.UtcNow.Date;
                    result.Add(new BulkChangeAssociateCustomerModel()
                    {
                        ServiceProviderId = serviceProviderId,
                        RevCustomerId = string.IsNullOrWhiteSpace(line.RevioCustomer) ? string.Empty : line.RevioCustomer,
                        ServiceTypeId = (int)DeviceChangeType.CustomerAssignment,
                        Devices = !string.IsNullOrWhiteSpace(line.ICCID) ? new string[] { line.ICCID } : new string[] { },
                        Prorate = !string.IsNullOrWhiteSpace(line.IsProrate) && line?.IsProrate.ToLower() == "y",
                        RevProductIdList = !string.IsNullOrWhiteSpace(line.RevioProduct) ? new List<int?> { int.Parse(line.RevioProduct) } : new List<int?>(),
                        Description = string.IsNullOrWhiteSpace(line.Description) ? string.Empty : line.Description,
                        RateList = !string.IsNullOrWhiteSpace(line.Rate) ? new List<decimal?>() { decimal.Parse(line.Rate) } : new List<decimal?>(),
                        IsUseCarrierActivation = !string.IsNullOrWhiteSpace(line.IsUseCarrierActivation) && line?.IsUseCarrierActivation.ToLower() == "y",
                        EffectiveDate = effectiveDate,
                        ActivatedDate = effectiveDate,
                        CreateRevService = !string.IsNullOrWhiteSpace(line.IsCreateServiceProduct) && line?.IsCreateServiceProduct.ToLower() == "y",
                        AddCustomerRatePlan = !string.IsNullOrWhiteSpace(line.IsAddCustomerRatePlan) && line?.IsAddCustomerRatePlan.ToLower() == "y",
                        CustomerRatePlan = string.IsNullOrWhiteSpace(line.CustomerRatePlanCode) ? string.Empty : line.CustomerRatePlanCode,
                        CustomerRatePool = string.IsNullOrWhiteSpace(line.CustomerPool) ? string.Empty : line.CustomerPool
                    });
                }
            }

            return result;
        }

        private List<ValidateExcelUpload> ValidateActivationInformationChange(int serviceProviderId, IList<ActivationInformationChangeCsvRow> lines)
        {
            var errorList = new List<ValidateExcelUpload>();

            var foundationAccList = altaWrxDb.usp_Telegence_Get_BillingAccounts().ToList();

            var jasperCarrierRatePlanRepository = new JasperCarrierRatePlanRepository(altaWrxDb);
            var ratePlans = jasperCarrierRatePlanRepository.GetAllByServiceProvider(serviceProviderId);
            var jasperCarrierRatePlans = jasperCarrierRatePlanRepository.GetAllByServiceProvider(serviceProviderId);

            var mobilityFeatureRepository = new MobilityFeatureRepository(altaWrxDb);
            var mobilityFeatures = mobilityFeatureRepository.GetAllByServiceProvider(serviceProviderId);

            var billingPeriodRepository = new BillingPeriodRepository(altaWrxDb);
            var currentBillingPeriodStartDate = billingPeriodRepository.GetCurrentBillingPeriodByServiceProviderId(serviceProviderId);

            var mobilityDeviceUsageAggregateRepository = new MobilityDeviceUsageAggregateRepository(altaWrxDb, serviceProviderId);
            var dataGroups = mobilityDeviceUsageAggregateRepository.GetGroups();
            var dataPools = mobilityDeviceUsageAggregateRepository.GetPools();

            var customerRatePlanRepository = new JasperCustomerRatePlanRepository(altaWrxDb);
            var customerRatePlans = customerRatePlanRepository.GetEligibleCustomerRatePlanCodes(serviceProviderId, permissionManager.Tenant.id, permissionManager);

            var qualificationTokens = lines.Select(x => x.QualificationToken).Distinct().ToList();
            var qualifiedAddresses = altaWrxDb.QualificationAddresses.Where(x => qualificationTokens.Contains(x.QualificationToken)
                    && x.IsQualified && !string.IsNullOrEmpty(x.QualificationToken) && !x.IsActivatedByService && !x.IsDeleted && x.IsActive).ToList();

            //input required: account number; rate plan code; first name; last name; street number; street name; city; state; zip code;
            foreach (var line in lines.Select((value, index) => new { index, value }))
            {
                if (!permissionManager.IsAgent)
                {
                    if (string.IsNullOrEmpty(line.value.BillingAccountNumber))
                    {
                        errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.BillingAccountNumberField, Resources.CommonStrings.CannotEmpty));
                    }
                    else if (!foundationAccList.Any(acc => acc.BillingAccountNumber == line.value.BillingAccountNumber))
                    {
                        errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.BillingAccountNumberField, Resources.CommonStrings.Invalid));
                    }

                    if (string.IsNullOrEmpty(line.value.CarrierRatePlanCode))
                    {
                        errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.CarrierRatePlanCodeField, Resources.CommonStrings.CannotEmpty));
                    }
                    else if (!ratePlans.Any(md => md.RatePlanCode == line.value.CarrierRatePlanCode))
                    {
                        errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.CarrierRatePlanCodeField, Resources.CommonStrings.Invalid));
                    }
                }
                else
                {
                    if (!string.IsNullOrWhiteSpace(line.value.CustomerRatePlanCode))
                    {
                        if (!customerRatePlans.Any(rl => rl.RatePlanCode == line.value.CustomerRatePlanCode))
                        {
                            errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.CustomerRatePlanCodeField, Resources.CommonStrings.Invalid));
                        }
                    }
                }

                if (!string.IsNullOrEmpty(line.value.AddSOC))
                {
                    var SOCCodes = line.value.AddSOC.Split(',');
                    var SOCCodeErrors = new List<string>();
                    for (int i = 0; i < SOCCodes.Length; i++)
                    {
                        if (!mobilityFeatures.Any(x => x.SOCCode == SOCCodes[i]))
                        {
                            SOCCodeErrors.Add(SOCCodes[i]);
                        }
                    }
                    if (SOCCodeErrors.Count > 0)
                    {
                        errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.RatePlanSOCField, string.Join(",", SOCCodeErrors) + " " + Resources.CommonStrings.Invalid));
                    }
                }

                if (!string.IsNullOrEmpty(line.value.RemoveSOC))
                {
                    var SOCCodes = line.value.RemoveSOC.Split(',');
                    var SOCCodeErrors = new List<string>();
                    for (int i = 0; i < SOCCodes.Length; i++)
                    {
                        if (!mobilityFeatures.Any(x => x.SOCCode == SOCCodes[i]))
                        {
                            SOCCodeErrors.Add(SOCCodes[i]);
                        }
                    }
                    if (SOCCodeErrors.Count > 0)
                    {
                        errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.RatePlanSOCField, string.Join(",", SOCCodeErrors) + " " + Resources.CommonStrings.Invalid));
                    }
                }

                if (!string.IsNullOrEmpty(line.value.CarrierRatePlanGroup) && !dataGroups.Any(gr => gr.DataGroupId == line.value.CarrierRatePlanGroup))
                {
                    errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.CarrierRatePlanGroupField, Resources.CommonStrings.Invalid));
                }

                if (!string.IsNullOrEmpty(line.value.CarrierRatePool) && !dataPools.Any(gr => gr.PoolId == line.value.CarrierRatePool))
                {
                    errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.CarrierRatePoolField, Resources.CommonStrings.Invalid));
                }

                if (string.IsNullOrEmpty(line.value.SubscriberFirstName))
                {
                    errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.SubscriberFirstNameField, Resources.CommonStrings.CannotEmpty));
                }

                if (string.IsNullOrEmpty(line.value.SubscriberLastName))
                {
                    errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.SubscriberLastNameField, Resources.CommonStrings.CannotEmpty));
                }

                if (string.IsNullOrEmpty(line.value.StreetNumber))
                {
                    errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.StreetNumberFieldField, Resources.CommonStrings.CannotEmpty));
                }

                if (string.IsNullOrEmpty(line.value.StreetName))
                {
                    errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.StreetNameField, Resources.CommonStrings.CannotEmpty));
                }

                if (string.IsNullOrEmpty(line.value.City))
                {
                    errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.CityField, Resources.CommonStrings.CannotEmpty));
                }
                if (string.IsNullOrEmpty(line.value.State))
                {
                    errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.StateField, Resources.CommonStrings.CannotEmpty));
                }

                if (string.IsNullOrEmpty(line.value.Zip))
                {
                    errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.ZipField, Resources.CommonStrings.CannotEmpty));
                }

                if (string.IsNullOrEmpty(line.value.SIM_ICCID))
                {
                    errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.SIM_ICCIDField, Resources.CommonStrings.CannotEmpty));
                }

                if (string.IsNullOrEmpty(line.value.IMEI))
                {
                    errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.IMEIField, Resources.CommonStrings.CannotEmpty));
                }

                // validate qualification addresses
                if (!string.IsNullOrEmpty(line.value.QualificationToken))
                {
                    var address = qualifiedAddresses.FirstOrDefault(x => x.QualificationToken == line.value.QualificationToken
                        && x.StreetNumber == line.value.StreetNumber
                        && x.StreetName == line.value.StreetName
                        && x.City == line.value.City
                        && x.State == line.value.State
                        && x.ZipCode == line.value.Zip);

                    if (address == null)
                    {
                        errorList.Add(new ValidateExcelUpload(line.index + 1, "Address", Resources.CommonStrings.AddressNotQualified));
                    }
                }
            }

            return errorList;
        }

        private List<ValidateExcelUpload> ValidateAssignCustomer(int serviceProviderId, IList<AssignCustomerCsvRow> lines)
        {
            var errorList = new List<ValidateExcelUpload>();

            var revCustomerList = ListHelper.GetRevCustomerList(permissionManager, permissionManager.Tenant.id, permissionManager.PermissionFilter.GetRevAccountFilter());
            var serviceTypeList = ListHelper.GetServiceTypeList(permissionManager);
            var customerRatePlanRepository = new JasperCustomerRatePlanRepository(altaWrxDb);
            var customerRatePlans = customerRatePlanRepository.GetEligibleCustomerRatePlanCodes(serviceProviderId, permissionManager.Tenant.id, permissionManager);
            var CustomerRatePools = altaWrxDb.CustomerRatePools
                .Where(ratePool => ratePool.IsActive
                                   && !ratePool.IsDeleted
                                   && ratePool.TenantId == permissionManager.Tenant.id)
                .ToList();

            foreach (var line in lines.Select((value, index) => new { index, value }))
            {
                if (string.IsNullOrWhiteSpace(line.value.RevioCustomer))
                {
                    errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.RevioCustomerField, Resources.CommonStrings.CannotEmpty));
                }
                else if (!revCustomerList.Any(acc => acc.RevCustomerId == line.value.RevioCustomer))
                {
                    errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.RevioCustomerField, Resources.CommonStrings.Invalid));
                }

                if (!string.IsNullOrWhiteSpace(line.value.IsCreateServiceProduct) && line.value.IsCreateServiceProduct.ToLower() == "y")
                {
                    // service type
                    if (string.IsNullOrWhiteSpace(line.value.ServiceType))
                    {
                        errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.RevioProductField, Resources.CommonStrings.CannotEmpty));
                    }
                    else if (!serviceTypeList.Any(acc => acc.ServiceTypeId == int.Parse(line.value.ServiceType)))
                    {
                        errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.RevioProductField, Resources.CommonStrings.Invalid));
                    }

                    var revIOProductList = ListHelper.GetRevProductList(permissionManager, line.value.RevioCustomer);
                    // rev product
                    if (string.IsNullOrWhiteSpace(line.value.RevioProduct))
                    {
                        errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.RevioProductField, Resources.CommonStrings.CannotEmpty));
                    }
                    else if (!revIOProductList.Any(acc => acc.ProductId == int.Parse(line.value.RevioProduct)))
                    {
                        errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.RevioProductField, Resources.CommonStrings.Invalid));
                    }

                    // rate
                    if (string.IsNullOrWhiteSpace(line.value.Rate))
                    {
                        errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.ServiceProductRateField, Resources.CommonStrings.CannotEmpty));
                    }
                }

                if (!string.IsNullOrWhiteSpace(line.value.IsProrate) && line.value.IsProrate.ToLower() != "y")
                {
                    if (string.IsNullOrWhiteSpace(line.value.IsUseCarrierActivation) && string.IsNullOrEmpty(line.value.EffectiveDate))
                    {
                        errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.IsUseCarrierActivationField, Resources.CommonStrings.CannotEmpty));
                    }
                }

                //customer rate plan
                if (!string.IsNullOrWhiteSpace(line.value.IsAddCustomerRatePlan) && line.value.IsAddCustomerRatePlan.ToLower() == "y")
                {
                    if (!string.IsNullOrWhiteSpace(line.value.CustomerRatePlanCode))
                    {
                        if (!customerRatePlans.Any(rl => rl.RatePlanCode == line.value.CustomerRatePlanCode))
                        {
                            errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.CustomerRatePlanCodeField, Resources.CommonStrings.Invalid));
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(line.value.CustomerPool))
                    {
                        if (!CustomerRatePools.Any(rp => rp.id == int.Parse(line.value.CustomerPool)))
                        {
                            errorList.Add(new ValidateExcelUpload(line.index + 1, Resources.CommonStrings.CustomerPoolField, Resources.CommonStrings.Invalid));
                        }
                    }
                }
            }

            return errorList;
        }


        public bool ValidateCarrierRatePlanEffectiveDate(DateTime effectiveDate, DateTime currentBillingPeriodStartDate)
        {
            return effectiveDate >= currentBillingPeriodStartDate && effectiveDate < DateTime.Now.AddDays(30);
        }

        private async Task<string> ProcessCarrierRatePlanChange(int serviceProviderId, int integrationId, IEnumerable<CarrierRatePlanChangeCsvRow> lines,
            IList<Models.Repositories.MobilityDevice> mobilityDevices)
        {
            var changeRepository = new DeviceBulkChangeRepository(altaWrxDb, permissionManager);
            var changeType = DeviceChangeType.CarrierRatePlanChange;
            var createdBy = SessionHelper.GetAuditByName(Session);

            var deviceChanges = new List<Mobility_DeviceChange>();
            foreach (var line in lines)
            {
                var mobilityDevice = mobilityDevices.FirstOrDefault(md => md.MSISDN == line.PhoneNumber);
                deviceChanges.Add(new Mobility_DeviceChange(CreateCarrierRatePlanChangeRequest(line, serviceProviderId, integrationId, mobilityDevice?.IMEI), mobilityDevice?.id, line.ICCID, line.PhoneNumber, createdBy));
            }

            var bulkChange = new DeviceBulkChange
            {
                ChangeRequestTypeId = (int)changeType,
                ServiceProviderId = serviceProviderId,
                TenantId = permissionManager.PermissionFilter.LoggedInTenantId,
                SiteId = GetSiteIdForBulkChange(deviceChanges),
                Status = deviceChanges.Any(change => !change.IsProcessed) ? BulkChangeStatus.NEW : BulkChangeStatus.PROCESSED,
                CreatedDate = DateTime.UtcNow,
                CreatedBy = SessionHelper.GetAuditByName(Session),
                IsActive = true,
                IsDeleted = false,
                Mobility_DeviceChange = deviceChanges
            };
            changeRepository.CreateBulkChange(bulkChange);

            var bulkChangeId = bulkChange.id;

            var customObjectDbList = GetTenantCustomFields();
            var awsAccessKey = AwsAccessKeyFromCustomObjects(customObjectDbList);
            var awsSecretAccessKey = AwsSecretAccessKeyFromCustomObjects(customObjectDbList);
            var queueName = ValueFromCustomObjects(customObjectDbList, "Device Bulk Change Queue");
            var sqsHelper = new SqsHelper(awsAccessKey, awsSecretAccessKey);
            var errorMessage = await sqsHelper.EnqueueBulkChangeAsync(queueName, bulkChangeId);
            if (string.IsNullOrEmpty(errorMessage))
                await UpdateBulkChange(bulkChangeId);

            return !string.IsNullOrWhiteSpace(errorMessage) ? $"An error occurred: {errorMessage}" : string.Empty;
        }

        private async Task UpdateBulkChange(long bulkChangeId)
        {
            var bulkChangeRepository = new DeviceBulkChangeRepository(altaWrxDb, permissionManager);
            var bulkChange = bulkChangeRepository.GetDeviceBulkChange(bulkChangeId);
            var processedBy = SessionHelper.GetAuditByName(Session);
            var processedDate = DateTime.UtcNow;
            bulkChange.Status = BulkChangeStatus.PROCESSING;
            bulkChange.ProcessedBy = processedBy;
            bulkChange.ProcessedDate = processedDate;
            bulkChange.ModifiedBy = processedBy;
            bulkChange.ModifiedDate = processedDate;
            await altaWrxDb.SaveChangesAsync();
        }

        private string CreateCarrierRatePlanChangeRequest(CarrierRatePlanChangeCsvRow line, int serviceProviderId, int integrationId, string imei)
        {
            return JsonConvert.SerializeObject(line.ToCarrierRatePlanChangeRequest(serviceProviderId, integrationId, imei));
        }

        private int? GetSiteIdForBulkChange(IEnumerable<Mobility_DeviceChange> changes)
        {
            if (permissionManager.IsParentAdmin)
            {
                return null;
            }

            var userSiteId = UserSiteId;
            if (userSiteId != null)
            {
                return userSiteId;
            }

            var deviceIds = changes.Select(change => change.DeviceId).ToList();
            return altaWrxDb.MobilityDevice_Tenant
                .FirstOrDefault(device => device.TenantId == permissionManager.Tenant.id && deviceIds.Contains(device.id) && device.SiteId != null)?.SiteId;
        }

        private IEnumerable<Mobility_DeviceChange> BuildChangeDetails(BulkChangeCreateModel bulkChange, int serviceProviderId, DeviceChangeType changeType)
        {
            switch (changeType)
            {
                case DeviceChangeType.StatusUpdate:
                    return BuildStatusUpdateChangeDetails(bulkChange, serviceProviderId);
                case DeviceChangeType.ActivateNewService:
                    return BuildNewServiceActivationTelegence(
                        Session,
                        altaWrxDb,
                        bulkChange.Devices.Where(d => !string.IsNullOrWhiteSpace(d)).ToList(),
                        bulkChange.StatusUpdate, bulkChange.CompatibleRatePlan, bulkChange.QualificationToken);
                case DeviceChangeType.Archival:
                    return BuildArchivalChangeDetails(bulkChange, serviceProviderId);
                case DeviceChangeType.CarrierRatePlanChange:
                    return BuildCarrierRatePlanChangeDetails(bulkChange, serviceProviderId);
                case DeviceChangeType.CustomerRatePlanChange:
                    return BuildCustomerRatePlanChangeDetails(bulkChange, serviceProviderId);
                case DeviceChangeType.EditUsername:
                    return BuildUsernameChangeDetails(bulkChange, serviceProviderId);
                default:
                    throw new NotImplementedException($"Unsupported device change type: {changeType}");
            }
        }

        private IEnumerable<Mobility_DeviceChange> BuildCustomerRatePlanChangeDetails(BulkChangeCreateModel bulkChange, int serviceProviderId)
        {
            var numbers = bulkChange.Devices.Where(number => !string.IsNullOrWhiteSpace(number)).ToList();
            var devicesByNumber = GetDevicesByNumber(altaWrxDb, serviceProviderId, numbers);
            var createdBy = SessionHelper.GetAuditByName(Session);

            var deviceChanges = new List<Mobility_DeviceChange>();
            var changeRequest = new DeviceChangeRequest(JsonConvert.SerializeObject(bulkChange,
                    new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore }), DateTime.UtcNow, createdBy);
            foreach (var number in numbers)
            {
                if (!devicesByNumber.ContainsKey(number))
                {
                    deviceChanges.Add(CreateDeviceChangeError(number, "Invalid subscriber number", createdBy));
                }
                else
                {
                    var device = devicesByNumber[number];
                    var deviceChange = GetMobilityDeviceChange(changeRequest, device, number, createdBy);
                    deviceChanges.Add(deviceChange);
                }
            }

            return deviceChanges;
        }

        private IEnumerable<Mobility_DeviceChange> BuildCarrierRatePlanChangeDetails(BulkChangeCreateModel bulkChange, int serviceProviderId)
        {
            var numbers = bulkChange.Devices.Where(number => !string.IsNullOrWhiteSpace(number)).ToList();
            var createdBy = SessionHelper.GetAuditByName(Session);

            var deviceChanges = new List<Mobility_DeviceChange>();
            var mobilityDeviceRepository = new MobilityDeviceRepository(altaWrxDb);
            var mobilityDevices = mobilityDeviceRepository.GetAllByServiceProviderId(serviceProviderId);
            foreach (var number in numbers)
            {
                var mobilityDevice = mobilityDevices.FirstOrDefault(md => md.MSISDN == number);
                if (mobilityDevice != null)
                {
                    var carrierRatePlanChange = new CarrierRatePlanChangeCsvRow
                    {
                        ICCID = mobilityDevice.ICCID,
                        PhoneNumber = number,
                        RatePlanName = bulkChange.CarrierRatePlanUpdate.CarrierRatePlan,
                        OptimizationGroup = bulkChange.CarrierRatePlanUpdate.OptimizationGroup,
                        EffectiveDate = bulkChange.CarrierRatePlanUpdate.EffectiveDate
                    };
                    deviceChanges.Add(new Mobility_DeviceChange(CreateCarrierRatePlanChangeRequest(carrierRatePlanChange, serviceProviderId, (int)bulkChange.IntegrationId, mobilityDevice.IMEI), mobilityDevice.id, carrierRatePlanChange.ICCID, carrierRatePlanChange.PhoneNumber, createdBy));
                }
                else
                {
                    deviceChanges.Add(CreateDeviceChangeError(number, string.Format(CommonStrings.MobilityDeviceSubscriberNumberNotExistError, number), createdBy));
                }
            }

            return deviceChanges;
        }

        private IEnumerable<Mobility_DeviceChange> BuildUsernameChangeDetails(BulkChangeCreateModel bulkChange, int serviceProviderId)
        {
            var numbers = bulkChange.Devices.Where(number => !string.IsNullOrWhiteSpace(number)).ToList();
            var devicesByNumber = GetDevicesByNumber(altaWrxDb, serviceProviderId, numbers);
            var createdBy = SessionHelper.GetAuditByName(Session);
            var deviceChanges = new List<Mobility_DeviceChange>();

            var changeRequest = new DeviceChangeRequest(JsonConvert.SerializeObject(bulkChange.Username, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore }), DateTime.UtcNow, createdBy);
            foreach (var number in numbers)
            {
                if (!devicesByNumber.ContainsKey(number))
                {
                    deviceChanges.Add(CreateDeviceChangeError(number, "Invalid subscriber number", createdBy));
                }
                else
                {
                    var device = devicesByNumber[number];
                    deviceChanges.Add(new Mobility_DeviceChange(changeRequest, device.id, device.ICCID, device.MSISDN));
                }
            }

            return deviceChanges;
        }

        private IEnumerable<Mobility_DeviceChange> BuildArchivalChangeDetails(BulkChangeCreateModel bulkChange, int serviceProviderId)
        {
            var numbers = bulkChange.Devices.Where(number => !string.IsNullOrWhiteSpace(number)).ToList();
            var devicesByNumber = GetDevicesByNumber(altaWrxDb, serviceProviderId, numbers);
            var activeRevServicesByDeviceId = GetActiveRevServicesByDeviceId(devicesByNumber.Values.Select(device => device.id).ToList());
            var createdBy = SessionHelper.GetAuditByName(Session);
            var archivalRecentUsageCutoff = DateTime.UtcNow.Date.AddDays(-1 * ARCHIVAL_RECENT_USAGE_VALIDATION_DAYS);
            return GetArchivalChanges(bulkChange, numbers, devicesByNumber, activeRevServicesByDeviceId, createdBy, archivalRecentUsageCutoff);
        }

        private static IEnumerable<Mobility_DeviceChange> GetArchivalChanges(BulkChangeCreateModel bulkChange, IEnumerable<string> numbers, ConcurrentDictionary<string, Models.Repositories.MobilityDevice> devicesByNumber, ConcurrentDictionary<int, ConcurrentBag<RevService>> activeRevServicesByDeviceId, string createdBy, DateTime archivalRecentUsageCutoff)
        {
            var deviceChanges = new List<Mobility_DeviceChange>();

            var changeRequest = new DeviceChangeRequest(JsonConvert.SerializeObject(bulkChange,
                    new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore }), DateTime.UtcNow, createdBy);
            foreach (var number in numbers)
            {
                if (!devicesByNumber.ContainsKey(number))
                {
                    deviceChanges.Add(CreateDeviceChangeError(number, "Invalid subscriber number", createdBy));
                }
                else
                {
                    var device = devicesByNumber[number];
                    var deviceChange = GetArchivalMobilityDeviceChange(bulkChange, device, number, activeRevServicesByDeviceId, createdBy, archivalRecentUsageCutoff, changeRequest);
                    deviceChanges.Add(deviceChange);
                }
            }

            return deviceChanges;
        }

        private static Mobility_DeviceChange GetArchivalMobilityDeviceChange(BulkChangeCreateModel bulkChange, Models.Repositories.MobilityDevice device, string number,
            ConcurrentDictionary<int, ConcurrentBag<RevService>> activeRevServicesByDeviceId, string createdBy, DateTime archivalRecentUsageCutoff, DeviceChangeRequest changeRequest)
        {
            if (activeRevServicesByDeviceId.ContainsKey(device.id) && !bulkChange.OverrideValidation.GetValueOrDefault(false))
            {
                var activeRevServiceIds = activeRevServicesByDeviceId[device.id].Select(svc => svc.RevServiceId);
                var errorMessage = $"Active Rev Services found associated with device: {string.Join(",", activeRevServiceIds)}";
                return CreateDeviceChangeError(number, errorMessage, createdBy);
            }

            if (device.LastUsageDate.HasValue && device.LastUsageDate.Value > archivalRecentUsageCutoff)
            {
                var errorMessage =
                    $"Device has had usage in the last {ARCHIVAL_RECENT_USAGE_VALIDATION_DAYS} days and is ineligible to be archived";
                return CreateDeviceChangeError(number, errorMessage, createdBy);
            }

            return GetMobilityDeviceChange(changeRequest, device, number, createdBy);
        }

        private static Mobility_DeviceChange GetMobilityDeviceChange(DeviceChangeRequest changeRequest, MobilityDevice device, string number, string createdBy)
        {
            return new Mobility_DeviceChange(changeRequest, device.id, device.ICCID, number);
        }

        private IEnumerable<Mobility_DeviceChange> BuildStatusUpdateChangeDetails(BulkChangeCreateModel bulkChange, int serviceProviderId)
        {
            var serviceProviderRepository = new ServiceProviderRepository(altaWrxDb);
            var serviceProvider = serviceProviderRepository.GetById(serviceProviderId);
            var statusUpdate = bulkChange.StatusUpdate;
            var numbers = bulkChange.Devices.Where(number => !string.IsNullOrWhiteSpace(number)).ToList();
            var integrationType = (IntegrationType)serviceProvider.IntegrationId;
            switch (integrationType)
            {
                case IntegrationType.Telegence:
                    return BuildStatusUpdateChangeDetailsTelegence(serviceProviderId, numbers, statusUpdate, altaWrxDb, Session, permissionManager);
                case IntegrationType.eBonding:
                    return BuildStatusUpdateChangeDetailsEbonding(serviceProviderId, numbers, statusUpdate, altaWrxDb, Session, user);
                default:
                    return null;
            }

        }

        public static IEnumerable<Mobility_DeviceChange> BuildStatusUpdateChangeDetailsTelegence(int serviceProviderId, ICollection<string> numbers, BulkChangeStatusUpdate statusUpdate, AltaWorxCentral_Entities altaWrxDb, HttpSessionStateBase session, PermissionManager permissionManager)
        {
            var devicesByNumber = GetDevicesByNumber(altaWrxDb, serviceProviderId, numbers);
            var targetStatusEntity = altaWrxDb.DeviceStatus.First(x =>
                x.IsActive && !x.IsDeleted
                && (x.Status == statusUpdate.TargetStatus || x.Description == statusUpdate.TargetStatus)
                && x.IntegrationId == (int)IntegrationType.Telegence);
            var targetStatus = targetStatusEntity.Status;
            var targetStatusId = targetStatusEntity.id;
            var createdBy = SessionHelper.GetAuditByName(session);

            var deviceChanges = new List<Mobility_DeviceChange>();
            var isSameChangeRequest = targetStatus != DeviceStatusConstant.Telegence_Active && !statusUpdate.CreateServiceProduct;
            DeviceChangeRequest changeRequest = null;
            if (isSameChangeRequest)
            {
                changeRequest = new DeviceChangeRequest(JsonConvert.SerializeObject(BuildTelegenceCommonStatusUpdateRequest(targetStatus, targetStatusId, statusUpdate), new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore }), DateTime.UtcNow, createdBy);
            }
            foreach (var number in numbers)
            {
                if (!devicesByNumber.ContainsKey(number))
                {
                    deviceChanges.Add(CreateDeviceChangeError(number, string.Format(CommonStrings.MobilityDeviceSubscriberNumberNotExistError, number), createdBy));
                }
                else
                {
                    var device = devicesByNumber[number];
                    if (!isSameChangeRequest)
                    {
                        changeRequest = BuildDeviceSpecificTelegenceStatusUpdateRequest(targetStatus, targetStatusId, statusUpdate, createdBy, device, altaWrxDb, permissionManager);
                    }
                    if (changeRequest == null)
                    {
                        deviceChanges.Add(CreateDeviceChangeError(number, string.Format(CommonStrings.ErrorCreatingDeviceChangeRequestForSubscriberNumber, number), createdBy));
                    }
                    else
                    {
                        deviceChanges.Add(new Mobility_DeviceChange(changeRequest, device.id, device.ICCID, number));
                    }
                }
            }

            return deviceChanges;
        }

        private static DeviceChangeRequest BuildDeviceSpecificTelegenceStatusUpdateRequest(string targetStatus, int targetStatusId, BulkChangeStatusUpdate statusUpdate, string createdBy, MobilityDevice device,
            AltaWorxCentral_Entities altaWrxDb, PermissionManager permissionManager)
        {
            BulkChangeStatusUpdateRequest<object> deviceRequest;
            switch (targetStatus.ToUpperInvariant())
            {
                case DeviceStatusConstant.Telegence_Active:
                    deviceRequest = BuildTelegenceActivateStatusUpdateRequest(targetStatus, targetStatusId, statusUpdate, device);
                    break;
                default:
                    deviceRequest = BuildTelegenceCommonStatusUpdateRequest(targetStatus, targetStatusId, statusUpdate);
                    break;
            }
            if (statusUpdate.CreateServiceProduct)
            {
                deviceRequest.RevServiceProductCreateModel = statusUpdate.RevServiceProductModel;
                deviceRequest.RevServiceProductCreateModel.DeviceId = device.id;
                deviceRequest.IntegrationAuthenticationId = GetIntegrationAuthenticationId(statusUpdate.RevServiceProductModel.RevCustomerId, altaWrxDb, permissionManager);
            }
            var changeRequestString = JsonConvert.SerializeObject(deviceRequest, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
            return new DeviceChangeRequest(changeRequestString, DateTime.UtcNow, createdBy);
        }

        internal static IEnumerable<Mobility_DeviceChange> BuildNewServiceActivationTelegence(
            HttpSessionStateBase session,
            AltaWorxCentral_Entities altaWrxDb,
            ICollection<string> devices, BulkChangeStatusUpdate statusUpdate,
            List<DeviceBulkChangeCompatibleRatePlan> compatibleRatePlans = null,
            string qualificationToken = "")
        {
            var createdBy = SessionHelper.GetAuditByName(session);

            var deviceChanges = new List<Mobility_DeviceChange>();
            foreach (var device in devices)
            {
                var ids = Regex.Split(device.Trim(), @"[^\d]+").Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
                if (ids.Length != 2)
                {
                    deviceChanges.Add(CreateDeviceChangeError(null, $"Invalid input line: {device}", createdBy));
                    continue;
                }
                var iccid = ids[0];
                var imei = ids[1];

                if (string.IsNullOrEmpty(iccid) || string.IsNullOrEmpty(imei))
                    deviceChanges.Add(CreateDeviceChangeError(null, $"Invalid input line: {device}", createdBy));

                var socCodes = statusUpdate.TelegenceStatusUpdate.FeatureCodes != null ? statusUpdate.TelegenceStatusUpdate.FeatureCodes.ToArray() : null;
                if (compatibleRatePlans != null && compatibleRatePlans.Count > 0)
                {
                    var compatibleRatePlan = compatibleRatePlans.FirstOrDefault(x => x.Device.Contains(device));
                    if (compatibleRatePlan != null)
                    {
                        statusUpdate.TelegenceStatusUpdate.RatePlanCode = compatibleRatePlan.RatePlanCodeSelected;
                        statusUpdate.TelegenceStatusUpdate.CustomerRatePlanId = compatibleRatePlan.CustomerRatePlanId;
                    }
                }

                var address = altaWrxDb.QualificationAddresses.FirstOrDefault(x => x.QualificationToken == qualificationToken
                    && x.IsQualified && !string.IsNullOrEmpty(x.QualificationToken) && !x.IsActivatedByService && !x.IsDeleted && x.IsActive);
                var change = new Mobility_DeviceChange(BuildTelegenceNewServiceActivationChangeRequest(iccid, imei, statusUpdate.TelegenceStatusUpdate, addSOC: socCodes, qualificationToken: qualificationToken), null, iccid, string.Empty, createdBy);
                change.QualificationAddressId = address?.id;

                deviceChanges.Add(change);
            }
            return deviceChanges;
        }
        private IEnumerable<Mobility_DeviceChange> BuildNewServiceActivationTelegenceByExcelFile(IList<ActivationInformationChangeCsvRow> lines)
        {
            var createdBy = SessionHelper.GetAuditByName(Session);
            var deviceChanges = new List<Mobility_DeviceChange>();
            var jasperCarrierRatePlanRepo = new JasperCarrierRatePlanRepository(altaWrxDb);
            var customerRatePlanCodes = lines.Select(x => x.CustomerRatePlanCode).ToList();
            var customerCarrierRatePlanMapping = jasperCarrierRatePlanRepo.GetCarrierRatePlanByCustomerRatePlanCode(customerRatePlanCodes);

            var qualificationTokens = lines.Select(x => x.QualificationToken).Distinct().ToList();
            var addresses = altaWrxDb.QualificationAddresses.Where(x => qualificationTokens.Contains(x.QualificationToken)
                    && x.IsQualified && !string.IsNullOrEmpty(x.QualificationToken) && !x.IsActivatedByService && !x.IsDeleted && x.IsActive).ToList();
            foreach (var line in lines)
            {
                var telegenceStatusUpdate = ChangeActivationInformationChangeCsvRowToBulkChangeStatusUpdateTelegence(line, customerCarrierRatePlanMapping);

                var address = addresses.FirstOrDefault(x => x.QualificationToken == line.QualificationToken
                                && x.StreetNumber == line.StreetNumber
                                && x.StreetName == line.StreetName
                                && x.City == line.City
                                && x.State == line.State
                                && x.ZipCode == line.Zip);

                var qualificationToken = address?.QualificationToken ?? string.Empty;
                var change = new Mobility_DeviceChange(BuildTelegenceNewServiceActivationChangeRequest(line.SIM_ICCID, line.IMEI, telegenceStatusUpdate, addSOC: line.AddSOC?.Split(','), qualificationToken: qualificationToken), null, line.SIM_ICCID, null, createdBy);
                change.QualificationAddressId = address?.id;

                deviceChanges.Add(change);
            }
            return deviceChanges;
        }

        private BulkChangeStatusUpdateTelegence ChangeActivationInformationChangeCsvRowToBulkChangeStatusUpdateTelegence(ActivationInformationChangeCsvRow line, Dictionary<JasperCustomerRatePlan, JasperCarrierRatePlan> customerCarrierRatePlanMapping)
        {
            var bulkChangeStatusUpdateTelegence = new BulkChangeStatusUpdateTelegence
            {
                FirstName = line.SubscriberFirstName,
                LastName = line.SubscriberLastName,
                StreetNo = line.StreetNumber,
                StreetName = line.StreetName,
                StreetDirection = line.StreetDirection,
                City = line.City,
                State = line.State,
                ZipCode = line.Zip,
                IMEI = line.IMEI,
                ReasonCode = "",
                SubscriberNo = "",
                RatePlanCode = line.CarrierRatePlanCode,
                BillingAccountNumber = line.BillingAccountNumber,
                CarrierDataGroup = line.CarrierRatePlanGroup,
                CarrierRatePool = line.CarrierRatePool,
            };
            if (permissionManager.IsAgent)
            {
                var mapping = customerCarrierRatePlanMapping.FirstOrDefault(x => x.Key.RatePlanCode == line.CustomerRatePlanCode);
                bulkChangeStatusUpdateTelegence.RatePlanCode = mapping.Value.RatePlanCode;
                bulkChangeStatusUpdateTelegence.CustomerRatePlanId = mapping.Key.id;
            }
            return bulkChangeStatusUpdateTelegence;
        }

        private static string BuildTelegenceNewServiceActivationChangeRequest(string iccid, string imei, BulkChangeStatusUpdateTelegence statusUpdate,
            string[] addSOC = null, string[] removeSOC = null, string qualificationToken = "")
        {
            TelegenceActivationRequest request = new TelegenceActivationRequest();
            BillingAccount billingAccount = new BillingAccount();
            billingAccount.Id = statusUpdate.BillingAccountNumber;
            request.BillingAccount = billingAccount;

            RelatedParty party = new RelatedParty();
            party.Id = "";
            party.Role = RELATED_PARTY_ROLE;
            party.FirstName = statusUpdate.FirstName;
            party.LastName = statusUpdate.LastName;

            Address address = new Address();
            address.StreetNumber = statusUpdate.StreetNo;
            address.StreetDirection = statusUpdate.StreetDirection ?? string.Empty;
            address.StreetName = statusUpdate.StreetName;
            address.City = statusUpdate.City;
            address.State = statusUpdate.State;
            address.ZipCode = statusUpdate.ZipCode;
            address.StreetType = string.Empty;
            party.Address = address;

            List<RelatedParty> relatedParties = new List<RelatedParty>();
            relatedParties.Add(party);
            request.RelatedParty = relatedParties;

            List<ServiceCharacteristic> serviceCharacteristics = new List<ServiceCharacteristic>();

            ServiceCharacteristic singleUserCode = new ServiceCharacteristic();
            singleUserCode.Name = "singleUserCode";
            singleUserCode.Value = statusUpdate.RatePlanCode;
            serviceCharacteristics.Add(singleUserCode);

            ServiceCharacteristic serviceZipCode = new ServiceCharacteristic();
            serviceZipCode.Name = "serviceZipCode";
            serviceZipCode.Value = statusUpdate.ZipCode;
            serviceCharacteristics.Add(serviceZipCode);

            ServiceCharacteristic equipmentType = new ServiceCharacteristic();
            equipmentType.Name = "equipmentType";
            equipmentType.Value = "G";
            serviceCharacteristics.Add(equipmentType);

            ServiceCharacteristic deviceTechnologyType = new ServiceCharacteristic();
            deviceTechnologyType.Name = "deviceTechnologyType";
            deviceTechnologyType.Value = "UMTS";
            serviceCharacteristics.Add(deviceTechnologyType);

            ServiceCharacteristic IMEI = new ServiceCharacteristic();
            IMEI.Name = "IMEI";
            IMEI.Value = imei;
            serviceCharacteristics.Add(IMEI);

            ServiceCharacteristic SIM = new ServiceCharacteristic();
            SIM.Name = "sim";
            SIM.Value = iccid;
            serviceCharacteristics.Add(SIM);

            for (int i = 0; i < addSOC?.Length; i++)
            {
                ServiceCharacteristic OfferingCode = new ServiceCharacteristic();
                OfferingCode.Name = $"offeringCode{i + 1}";
                OfferingCode.Value = addSOC[i];
                serviceCharacteristics.Add(OfferingCode);
            }

            for (int i = 0; i < removeSOC?.Length; i++)
            {
                ServiceCharacteristic OfferingCode = new ServiceCharacteristic();
                OfferingCode.Name = REMOVE_SOCCODE_STRING;
                OfferingCode.Value = removeSOC[i];
                serviceCharacteristics.Add(OfferingCode);
            }

            Service service = new Service();
            service.ServiceCharacteristic = serviceCharacteristics;

            Category category = new Category();
            category.Id = "";
            category.Name = "";
            service.Category = category;

            ServiceSpecification serviceSpecification = new ServiceSpecification();
            serviceSpecification.Id = "12";
            serviceSpecification.Href = "Mobility Services";
            service.ServiceSpecification = serviceSpecification;

            ServiceQualification serviceQualification = new ServiceQualification();
            serviceQualification.Id = qualificationToken;
            service.ServiceQualification = serviceQualification;

            service.Name = "";
            request.Service = service;

            var telegenceActivationChangeRequest = new TelegenceActivationChangeRequest
            {
                TelegenceActivationRequest = request,
                CarrierDataGroup = statusUpdate.CarrierDataGroup,
                CarrierRatePool = statusUpdate.CarrierRatePool,
                CustomerRatePlan = statusUpdate.CustomerRatePlanId,
                CustomerRatePool = null,
                AddCustomerRatePlan = statusUpdate.CustomerRatePlanId != null && statusUpdate.CustomerRatePlanId > 0,
                StaticIPProvision = statusUpdate.TelegenceIPProvision
            };
            return JsonConvert.SerializeObject(telegenceActivationChangeRequest);
        }

        public static IEnumerable<Mobility_DeviceChange> BuildStatusUpdateChangeDetailsEbonding(int serviceProviderId, ICollection<string> numbers, BulkChangeStatusUpdate statusUpdate, AltaWorxCentral_Entities altaWrxDb, HttpSessionStateBase session, User user)
        {
            var devicesByNumber = GetDevicesByNumber(altaWrxDb, serviceProviderId, numbers);
            var targetStatus = statusUpdate.TargetStatus;
            var postApiStatus = eBondingHelper.GetEbondingTargetStatus(targetStatus);
            var postApiStatusId = altaWrxDb.DeviceStatus.First(x =>
                x.IsActive && !x.IsDeleted && x.Status == postApiStatus && x.IntegrationId == (int)IntegrationType.eBonding).id;
            var createdBy = SessionHelper.GetAuditByName(session);
            var deviceChanges = new List<Mobility_DeviceChange>();
            foreach (var number in numbers)
            {
                if (!devicesByNumber.ContainsKey(number))
                {
                    deviceChanges.Add(CreateDeviceChangeError(number, "Invalid subscriber number", createdBy));
                }
                else
                {
                    var device = devicesByNumber[number];
                    var changeRequest = BuildStatusUpdateRequestEbonding(device, targetStatus, postApiStatusId, statusUpdate.eBondingStatusUpdate, user);
                    if (changeRequest == null)
                    {
                        deviceChanges.Add(CreateDeviceChangeError(number, "Could not create change request", createdBy));
                    }
                    else
                    {
                        deviceChanges.Add(new Mobility_DeviceChange(changeRequest, device.id, device.ICCID, number, createdBy));
                    }
                }
            }

            return deviceChanges;
        }

        private static ConcurrentDictionary<string, Models.Repositories.MobilityDevice> GetDevicesByNumber(AltaWorxCentral_Entities awxDb, int serviceProviderId, ICollection<string> numbers)
        {
            var devicesByNumber = new ConcurrentDictionary<string, Models.Repositories.MobilityDevice>(awxDb.MobilityDevices
                .Include(device => device.ServiceProvider).AsNoTracking()
                .Where(device => device.IsActive
                                 && !device.IsDeleted
                                 && serviceProviderId == device.ServiceProviderId
                                 && numbers.Contains(device.MSISDN))
                .GroupBy(x => x.MSISDN)
                .Select(g => g.FirstOrDefault())
                .AsEnumerable()
                .ToDictionary(device => device.MSISDN, device => device));
            return devicesByNumber;
        }
        private static ConcurrentDictionary<string, Models.Repositories.MobilityDevice> GetDevicesByIccid(AltaWorxCentral_Entities awxDb, int serviceProviderId, ICollection<string> iccids)
        {
            var devicesByIccid = new ConcurrentDictionary<string, Models.Repositories.MobilityDevice>(awxDb.MobilityDevices
                .Include(device => device.ServiceProvider).AsNoTracking()
                .Where(device => device.IsActive
                                 && !device.IsDeleted
                                 && serviceProviderId == device.ServiceProviderId
                                 && iccids.Contains(device.ICCID))
                .GroupBy(x => x.ICCID)
                .Select(g => g.FirstOrDefault())
                .AsEnumerable()
                .ToDictionary(device => device.ICCID, device => device));
            return devicesByIccid;
        }

        private static IEnumerable<string> CheckDevicesArchived(AltaWorxCentral_Entities awxDb, int serviceProviderId, ICollection<string> numbers)
        {
            var devicesByNumber = awxDb.MobilityDevices
                .Include(device => device.ServiceProvider).AsNoTracking()
                .Where(device => !device.IsActive
                                 && device.IsDeleted
                                 && serviceProviderId == device.ServiceProviderId
                                 && numbers.Contains(device.MSISDN))
                .GroupBy(device => device.MSISDN)
                .Select(g => g.FirstOrDefault() == null ? string.Empty : g.FirstOrDefault().MSISDN)
                .ToList();
            return devicesByNumber;
        }
        private static IEnumerable<string> CheckDevicesArchivedByIccid(AltaWorxCentral_Entities awxDb, int serviceProviderId, ICollection<string> numbers)
        {
            var devicesByIccid = awxDb.MobilityDevices
                .Include(device => device.ServiceProvider).AsNoTracking()
                .Where(device => !device.IsActive
                                 && device.IsDeleted
                                 && serviceProviderId == device.ServiceProviderId
                                 && numbers.Contains(device.ICCID))
                .GroupBy(device => device.ICCID)
                .Select(g => g.FirstOrDefault() == null ? string.Empty : g.FirstOrDefault().ICCID)
                .ToList();
            return devicesByIccid;
        }

        private ConcurrentDictionary<int, ConcurrentBag<RevService>> GetActiveRevServicesByDeviceId(ICollection<int> deviceIds)
        {
            var repo = new RevServiceRepository(altaWrxDb, 0);
            var revServices = repo.GetActiveRevServicesForDevices(deviceIds, PORTAL_TYPE, permissionManager.Tenant.id);

            var revServicesByDeviceId = new ConcurrentDictionary<int, ConcurrentBag<RevService>>();
            foreach (var revService in revServices)
            {
                revServicesByDeviceId.GetOrAdd(revService.Key, s => new ConcurrentBag<RevService>()).Add(revService.Value);
            }

            return revServicesByDeviceId;
        }

        private static Mobility_DeviceChange CreateDeviceChangeError(string subscriberNumber, string errorMessage, string createdBy)
        {
            return new Mobility_DeviceChange
            {
                Status = BulkChangeStatus.ERROR,
                SubscriberNumber = subscriberNumber,
                StatusDetails = errorMessage,
                IsProcessed = true,
                HasErrors = true,
                CreatedDate = DateTime.UtcNow,
                CreatedBy = createdBy,
                IsActive = true,
                IsDeleted = false,
            };
        }
        private static Mobility_DeviceChange CreateDeviceChangeErrorByIccid(string iccid, string errorMessage, string createdBy)
        {
            return new Mobility_DeviceChange
            {
                Status = BulkChangeStatus.ERROR,
                ICCID = iccid,
                StatusDetails = errorMessage,
                IsProcessed = true,
                HasErrors = true,
                CreatedDate = DateTime.UtcNow,
                CreatedBy = createdBy,
                IsActive = true,
                IsDeleted = false,
            };
        }


        private static BulkChangeStatusUpdateRequest<object> BuildTelegenceActivateStatusUpdateRequest(string targetStatus,
            int postUpdateStatusId, BulkChangeStatusUpdate statusUpdate, MobilityDevice device = null)
        {
            var telegenceStatusUpdate = statusUpdate.TelegenceStatusUpdate;
            if (string.IsNullOrWhiteSpace(targetStatus) || telegenceStatusUpdate == null)
            {
                return null;
            }

            var request = new BulkChangeStatusUpdateRequest<object>
            {
                UpdateStatus = targetStatus,
                PostUpdateStatusId = postUpdateStatusId,
                Request = BuildActivationStatusUpdateRequestTelegence(device, telegenceStatusUpdate)
            };

            return request;
        }

        private static BulkChangeStatusUpdateRequest<object> BuildTelegenceCommonStatusUpdateRequest(string targetStatus,
            int postUpdateStatusId, BulkChangeStatusUpdate statusUpdate)
        {
            var telegenceStatusUpdate = statusUpdate.TelegenceStatusUpdate;
            if (string.IsNullOrWhiteSpace(targetStatus) || telegenceStatusUpdate == null)
            {
                return null;
            }

            var updateRequest = new TelegenceSubscriberUpdateRequest
            {
                Mode = targetStatus,
                ReasonCode = telegenceStatusUpdate.ReasonCode
            };

            if (DateTime.TryParse(telegenceStatusUpdate.EffectiveDate, out var effectiveDate))
            {
                updateRequest.EffectiveDate = effectiveDate.ToString(CommonConstants.AMOP_UTC_DAY_TIME_FORMAT);
            }

            var request = new BulkChangeStatusUpdateRequest<object>
            {
                UpdateStatus = targetStatus,
                PostUpdateStatusId = postUpdateStatusId,
                Request = updateRequest
            };

            return request;
        }

        private static int GetIntegrationAuthenticationId(string revCustomerId, AltaWorxCentral_Entities altaWrxDb, PermissionManager permissionManager, int? test = null)
        {
            if (test != null)
            {
                test = 1;
            }
            var revCustomerRepository = new RevCustomerRepository(altaWrxDb, permissionManager.Tenant.id);
            var revCustomer = revCustomerRepository.GetByRevCustomerId(revCustomerId);
            return revCustomer.IntegrationAuthenticationId.GetValueOrDefault(0);
        }

        private static TelegenceActivationRequest BuildActivationStatusUpdateRequestTelegence(Models.Repositories.MobilityDevice device, BulkChangeStatusUpdateTelegence additionalDetails)
        {
            var service = new Service
            {
                ServiceCharacteristic = new List<ServiceCharacteristic>
                {
                    new ServiceCharacteristic
                    {
                        Name = "singleUserCode",
                        Value = additionalDetails.RatePlanCode
                    },
                    new ServiceCharacteristic
                    {
                        Name = "serviceZipCode",
                        Value = additionalDetails.ZipCode
                    },
                    new ServiceCharacteristic
                    {
                        Name = "equipmentType",
                        Value = "G"
                    },
                    new ServiceCharacteristic
                    {
                        Name = "deviceTechnologyType",
                        Value = "UMTS"
                    },
                    new ServiceCharacteristic
                    {
                        Name = "IMEI",
                        Value = device.IMEI
                    },
                    new ServiceCharacteristic
                    {
                        Name = "sim",
                        Value = device.ICCID
                    }
                },
                ServiceSpecification = new ServiceSpecification
                {
                    Id = "",
                    Href = ""
                },
                Category = new Category
                {
                    Id = "",
                    Name = ""
                },
                Name = ""
            };

            var activationRequest = new TelegenceActivationRequest
            {
                BillingAccount = new BillingAccount
                {
                    Id = device.BillingAccountNumber
                },
                RelatedParty = new List<RelatedParty>
                {
                    new RelatedParty
                    {
                        Id = "",
                        Role = "",
                        FirstName = additionalDetails.FirstName,
                        LastName = additionalDetails.LastName,
                        Address = new Address
                        {
                            StreetNumber = additionalDetails.StreetNo,
                            StreetDirection = additionalDetails.StreetDirection,
                            StreetName = additionalDetails.StreetName,
                            City = additionalDetails.City,
                            State = additionalDetails.State,
                            ZipCode = additionalDetails.ZipCode
                        }
                    }
                },
                Service = service
            };
            return activationRequest;
        }

        private static string BuildStatusUpdateRequestEbonding(IDevice device, string targetStatus, int postUpdateStatusId, eBondingDeviceUpdateRequest statusUpdate, User user)
        {
            if (string.IsNullOrWhiteSpace(targetStatus) || statusUpdate == null)
            {
                return null;
            }

            statusUpdate.ChangeRequestType = targetStatus;

            var careOrderRequest = eBondingHelper.GetEbondingCareOrderRequest(statusUpdate, device, user);
            var request = new BulkChangeStatusUpdateRequest<CareOrderRequest1>
            {
                UpdateStatus = targetStatus,
                PostUpdateStatusId = postUpdateStatusId,
                Request = careOrderRequest
            };
            return JsonConvert.SerializeObject(request, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
        }

        [HttpPost]
        public async Task<ActionResult> ProcessBulkChange(long id, long additionBulkChangeId = 0)
        {
            if (!permissionManager.UserCanEdit(Session, ModuleEnum.Mobility))
            {
                return new HttpStatusCodeResult(HttpStatusCode.Forbidden);
            }

            var bulkChange = altaWrxDb.DeviceBulkChanges.Find(id);
            if (bulkChange == null)
            {
                return new HttpNotFoundResult();
            }
            var newBulkChangeStatus = BulkChangeStatus.PROCESSED;
            if (bulkChange.Mobility_DeviceChange.Any(change => !change.IsProcessed))
            {
                newBulkChangeStatus = BulkChangeStatus.PROCESSING;
                var customObjectDbList = GetTenantCustomFields();
                var awsAccessKey = AwsAccessKeyFromCustomObjects(customObjectDbList);
                var awsSecretAccessKey = AwsSecretAccessKeyFromCustomObjects(customObjectDbList);
                var queueName = ValueFromCustomObjects(customObjectDbList, CommonConstants.CUSTOM_OBJECT_BULK_CHANGE_QUEUE_KEY);
                var sqsHelper = new SqsHelper(awsAccessKey, awsSecretAccessKey);
                var errorMessage = await sqsHelper.EnqueueBulkChangeAsync(queueName, id, additionBulkChangeId);
                if (!string.IsNullOrWhiteSpace(errorMessage))
                {
                    return Json(new { errors = new[] { $"An error occurred: {errorMessage}" } }, JsonRequestBehavior.AllowGet);
                }
            }

            var processedBy = SessionHelper.GetAuditByName(Session);
            var processedDate = DateTime.UtcNow;
            bulkChange.Status = newBulkChangeStatus;
            bulkChange.ProcessedBy = processedBy;
            bulkChange.ProcessedDate = processedDate;
            bulkChange.ModifiedBy = processedBy;
            bulkChange.ModifiedDate = processedDate;
            altaWrxDb.SaveChanges();

            return Json(new { redirectUrl = Url.Action("BulkChange", "Mobility", new { id }), status = "OK" }, JsonRequestBehavior.AllowGet);
        }

        private MobilityConfigurationChangeQueue CreateMobilityConfigurationChangeQueue(int deviceId,
            int serviceProviderId, List<string> mobilityConfigurationIDsToAdd, List<string> mobilityConfigurationIDsToRemove, List<string> mobilityConfigurationsCurrent,
            MobilityConfigurationType configurationType, string effectiveDate, int telegenceDeviceId, int? optimizationGroup = null)
        {
            var mobilityConfigurations = string.Empty;
            if (mobilityConfigurationsCurrent != null && mobilityConfigurationsCurrent.Count > 0)
            {
                mobilityConfigurations = string.Join(",", mobilityConfigurationsCurrent);
            }
            var details = new MobilityConfiguration
            {
                MobilityConfigurationIDsToAdd = mobilityConfigurationIDsToAdd,
                MobilityConfigurationIDsToRemove = mobilityConfigurationIDsToRemove,
                MobilityConfigurationsCurrent = mobilityConfigurations,
                ConfigurationType = configurationType.ToString(),
                OptimizationGroup = optimizationGroup,
                EffectiveDate = !string.IsNullOrWhiteSpace(effectiveDate) ? effectiveDate : DateTime.UtcNow.ToString(CommonConstants.AMOP_DATE_FORMAT),
                telegenceDeviceId = telegenceDeviceId
            };

            return CreateMobilityConfigurationChangeQueue(deviceId, serviceProviderId, JsonConvert.SerializeObject(details, new JsonSerializerSettings
            {
                NullValueHandling = NullValueHandling.Ignore
            }));
        }

        private MobilityConfigurationChangeQueue CreateMobilityConfigurationChangeQueue(int deviceId, int serviceProviderId, string details)
        {
            return new MobilityConfigurationChangeQueue
            {
                MobilityDeviceId = deviceId,
                ServiceProviderId = serviceProviderId,
                IsProcessed = false,
                CreatedBy = SessionHelper.GetAuditByName(Session),
                CreatedDate = DateTime.UtcNow,
                MobilityConfigurationChangeDetails = details,
                IsActive = true,
                IsDeleted = false,
                TenantId = SessionHelper.User(Session).TenantId
            };
        }

        private async Task<string> EnqueueMobilityConfigurationChangeAsync(int mobilityConfigurationChangeId)
        {
            try
            {
                var awsAccessInfo = GetAWSAccessInfo();
                var awsCredentials = new Amazon.Runtime.BasicAWSCredentials(awsAccessInfo.AWSAccessKey, awsAccessInfo.AWSSecretAccessKey);
                using (var client = new AmazonSQSClient(awsCredentials, Amazon.RegionEndpoint.USEast1))
                {
                    var queueList = client.ListQueues(GetMobilityConfigurationQueueName());
                    if (queueList.HttpStatusCode == HttpStatusCode.OK && queueList.QueueUrls != null && queueList.QueueUrls.Count > 0)
                    {
                        var request = new SendMessageRequest
                        {
                            MessageAttributes = new Dictionary<string, MessageAttributeValue>
                            {
                                {
                                    "MobilityLineConfigurationQueueId",
                                    new MessageAttributeValue
                                    {
                                        DataType = "String",
                                        StringValue = mobilityConfigurationChangeId.ToString()
                                    }
                                }
                            },
                            MessageBody = "Not used",
                            QueueUrl = queueList.QueueUrls[0]
                        };

                        var response = await client.SendMessageAsync(request);
                        if ((int)response.HttpStatusCode < 200 || (int)response.HttpStatusCode > 299)
                        {
                            return $"Error enqueuing mobility configuration change: {response.HttpStatusCode:d} {response.HttpStatusCode:g}";
                        }

                        // success
                        return string.Empty;
                    }

                    return "Error enqueuing mobility configuration change: Queue not found";
                }
            }
            catch (Exception ex)
            {
                var message = $"Error Queuing Mobility Line Configuration update for {mobilityConfigurationChangeId}";
                Log.Error(message, ex);
                return message;
            }
        }

        [HttpPost]
        public async Task<ActionResult> UpdateTelegenceStatus(string iccid, string telegenceStatus, string firstName, string lastName, string streetNo, string streetName, string streetDirection, string city, string state, string zipCode, string ratePlanCode = null, string reasonCode = "", string effectiveDate = "")
        {
            try
            {
                var telegenceDevice = altaWrxDb.TelegenceDevices.FirstOrDefault(x => x.ICCID == iccid);
                if (telegenceDevice == null)
                {
                    return ErrorMessage($"Could not find device with ICCID {iccid}");
                }

                var tad = new TelegenceAdditionalDetails();
                if (telegenceStatus == TelegenceDeviceStatus.A.ToString())
                {
                    tad.BillingAccountNo = telegenceDevice.BillingAccountNumber;
                    tad.FirstName = firstName;
                    tad.LastName = lastName;
                    tad.StreetNo = streetNo;
                    tad.StreetName = streetName;
                    tad.StreetDirection = streetDirection;
                    tad.City = city;
                    tad.State = state;
                    tad.ZipCode = zipCode;
                    tad.IMEI = telegenceDevice.IMEI;
                }
                else
                {
                    tad.ReasonCode = reasonCode;
                    tad.EffectiveDate = !string.IsNullOrEmpty(effectiveDate)
                        ? DateTime.Parse(effectiveDate).ToString("yyyy-MM-dd") + "Z"
                        : string.Empty;
                    tad.SubscriberNo = telegenceDevice.SubscriberNumber;
                }
                var res = await UpdateStatus(iccid, telegenceStatus, ratePlanCode, zipCode, tad);
                SessionHelper.SetAlert(Session, "Successfully started device status upload.");
                SessionHelper.SetAlertType(Session, "success");
                return res;
            }
            catch (Exception ex)
            {
                return ErrorMessage($"Error queuing the status change. {ex.Message}");
            }
        }

        public ActionResult TelegenceDeviceStatusReasonCodeDropdown(string status)
        {
            var model = ListHelper.TelegenceDeviceStatusReasonCodeList(permissionManager, status);
            return PartialView(model);
        }

        public FileContentResult TelegenceReasonCodeList()
        {
            var telegenceAllowedStatusList = altaWrxDb.DeviceStatus
                .Where(x => !x.IsDeleted && x.IsActive && x.AllowsApiUpdate && x.IntegrationId == (int)IntegrationType.Telegence)
                .OrderBy(x => x.Status)
                .ToList();

            var dsReasonCodes = new DataSet();
            foreach (var status in telegenceAllowedStatusList)
            {
                var telegenceDeviceStatusReasonCodes = altaWrxDb.TelegenceDeviceStatusReasonCodes
                    .Where(x => x.IsActive && !x.IsDeleted && x.DeviceStatu.Status == status.Status && x.DeviceStatu.IntegrationId == (int)IntegrationType.Telegence)
                    .ToList();

                if (telegenceDeviceStatusReasonCodes.Count > 0)
                {
                    // add to dataset
                    var tempDataSet = telegenceDeviceStatusReasonCodes.Select(x => new TelegenceStatusReasonCodeLite(x)).ToDataSet(status.DisplayName + " Codes");
                    dsReasonCodes.Merge(tempDataSet.Tables[0]);
                }
            }

            var bytes = ExcelUtilities.Export(dsReasonCodes);

            var file = File(bytes, ExcelContentType, $"TelegenceReasonCodes.{ExcelFileExtension}");
            return file;
        }

        public ActionResult NumberPortIn(int? serviceProviderId, string filter = "")
        {
            ViewBag.PageTitle = "Number Port-In";

            if (permissionManager.UserCannotAccess(Session, ModuleEnum.Mobility))
                return RedirectToAction("Index", "Home");

            var model = new NumberPortInModel(altaWrxDb, permissionManager.Tenant.id, serviceProviderId, filter);
            return View(model);
        }

        [HttpPost]
        public ActionResult TelegenceCheckEligibility(string subscriberNo, string zipCode)
        {
            try
            {
                TelegenceCheckEligibilityRequest request = new TelegenceCheckEligibilityRequest()
                {
                    SubscriberNumber = subscriberNo,
                    ServiceZipCode = zipCode
                };

                NumberPortInRepository NPIR = new NumberPortInRepository(altaWrxDb);
                var numberPortInRequest = NPIR.GetBySubscriberNo(subscriberNo);

                if (numberPortInRequest != null)
                {
                    return new JsonResult { Data = "Duplicate" };
                }
                else
                {
                    var response = TelegenceAPI.CheckPortInEligibility(request, altaWrxDb, permissionManager.Tenant.id);

                    var portInResponse =
                        JsonConvert.DeserializeObject<TelegenceCheckEligibilityResponse>(response.Content.ReadAsStringAsync().Result);

                    if (portInResponse.EligibilityFlag == "true")
                        return new JsonResult { Data = portInResponse.FromService.ServiceArea };
                    else
                        return new JsonResult { Data = "Error" };
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }

        public ActionResult TelegenceRequestPortIn(string subscriberNo, string zipCode, int serviceProviderId, string serviceArea = "")
        {
            ViewBag.PageTitle = "Telegence Port-In Request";

            var model = new TelegencePortInRequestModel(subscriberNo, zipCode, altaWrxDb, permissionManager, serviceArea, serviceProviderId);

            return View(model);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult TelegenceRequestPortIn(TelegencePortInRequestModel model)
        {
            if (ModelState.IsValid)
            {
                try
                {
                    var billingAccount = new BillingAccount
                    {
                        Id = model.BillingAccountId
                    };

                    var relatedParties = new List<RelatedParty>();
                    var relatedPartyAddressSubscriber = new Address
                    {
                        StreetNumber = model.RelatedPartyStreetNo_Subscriber,
                        StreetName = model.RelatedPartyStreetName_Subscriber,
                        StreetDirection = model.RelatedPartyStreetDirection_Subscriber,
                        StreetType = model.RelatedPartyStreetType_Subscriber,
                        City = model.RelatedPartyCity_Subscriber,
                        State = model.RelatedPartyState_Subscriber,
                        ZipCode = model.RelatedPartyZipCode_Subscriber
                    };
                    var partySubscriber = new RelatedParty
                    {
                        Id = "",
                        Role = "subscriber",
                        FirstName = model.RelatedPartyFirstName_Subscriber,
                        LastName = model.RelatedPartyLastName_Subscriber,
                        Address = relatedPartyAddressSubscriber
                    };
                    relatedParties.Add(partySubscriber);

                    var relatedPartyAddressAuthorizer = new Address
                    {
                        StreetNumber = model.RelatedPartyStreetNo_Authorizer,
                        StreetName = model.RelatedPartyStreetName_Authorizer,
                        StreetDirection = model.RelatedPartyStreetDirection_Authorizer,
                        StreetType = model.RelatedPartyStreetType_Authorizer,
                        City = model.RelatedPartyCity_Authorizer,
                        State = model.RelatedPartyState_Authorizer,
                        ZipCode = model.RelatedPartyZipCode_Authorizer
                    };
                    var partyAuthorizer = new RelatedParty
                    {
                        Id = "",
                        Role = "authorizer",
                        FirstName = model.RelatedPartyFirstName_Authorizer,
                        LastName = model.RelatedPartyLastName_Authorizer,
                        Address = relatedPartyAddressAuthorizer
                    };
                    relatedParties.Add(partyAuthorizer);

                    var fromService = new FromService
                    {
                        NpaNxx = model.SubscriberNo.Substring(0, 6),
                        FromLine = model.SubscriberNo.Substring(6, 4),
                        ServiceArea = model.FromService_ServiceArea
                    };

                    var oldServiceProvider = new OldServiceProvider
                    {
                        LocalId = "",
                        NetworkId = "",
                        BillingAccountNumber = model.OldServiceProviderBillingAccountNo,
                        FirstName = model.OldServiceProviderFirstName,
                        LastName = model.OldServiceProviderLastName,
                        BillingAccountPassword = "",
                        AuthorizationName = ""
                    };

                    var request = new TelegencePortInRequest
                    {
                        BillingAccount = billingAccount,
                        RelatedParty = relatedParties,
                        FromService = fromService,
                        OldServiceProvider = oldServiceProvider,
                        SubscriberNumber = model.SubscriberNo,
                        ServiceZipCode = model.ZipCode,
                        BusinessName = model.BusinessName,
                        PortDirection = "B",
                        PortRequestLineId = model.SubscriberNo
                    };

                    var result = TelegenceAPI.SendPortInRequest(request, altaWrxDb, permissionManager.Tenant.id, model.ServiceProviderId);
                    if (result != null)
                    {
                        var portInRepository = new NumberPortInRepository(altaWrxDb);
                        var portInRequest = new NumberPortInRequest
                        {
                            SubscriberNo = model.SubscriberNo,
                            ZipCode = model.ZipCode,
                            ServiceProviderId = model.ServiceProviderId,
                            PortInRequestStatus = altaWrxDb.TelegencePortInRequestStatus.FirstOrDefault(x => x.Status.ToLower() == result.portRequestStatus.ToLower())?.id,
                            RequestPayload = JsonConvert.SerializeObject(request),
                            ResponsePayload = JsonConvert.SerializeObject(result),
                            CreatedBy = user.Name,
                            CreatedDate = DateTime.UtcNow,
                            IsActive = true,
                            IsDeleted = false,
                            TenantId = permissionManager.Tenant.id
                        };
                        portInRepository.SaveNew(Session, portInRequest);
                        SessionHelper.SetAlert(Session, "Successfully sent number port in request. You may check request status below.");
                        SessionHelper.SetAlertType(Session, "success");
                        return RedirectToAction("NumberPortIn");
                    }
                    else
                    {
                        var msg = "Error sending Port-In request. Please enter valid data.";
                        TempData["Alert"] = msg;
                        TempData["AlertType"] = "danger";

                        model = new TelegencePortInRequestModel(model.SubscriberNo, model.ZipCode, altaWrxDb, permissionManager, model.FromService_ServiceArea, model.ServiceProviderId);
                        return View(model);
                    }
                }
                catch (Exception e)
                {
                    var msg = $"Exception: {e.Message}";
                    TempData["Alert"] = msg;
                    TempData["AlertType"] = "danger";

                    model = new TelegencePortInRequestModel(model.SubscriberNo, model.ZipCode, altaWrxDb, permissionManager, model.FromService_ServiceArea, model.ServiceProviderId);
                    return View(model);
                }
            }
            else
            {
                var msg = "Port-In request form is not valid. Please correct the errors before making a request.";
                TempData["Alert"] = msg;
                TempData["AlertType"] = "danger";

                model = new TelegencePortInRequestModel(model.SubscriberNo, model.ZipCode, altaWrxDb, permissionManager, model.FromService_ServiceArea, model.ServiceProviderId);
                return View(model);
            }
        }

        [HttpPost]
        public ActionResult CheckPortInRequestStatus(string subscriberNo)
        {
            try
            {
                var result = TelegenceAPI.NumberPortInStatusCheck(subscriberNo, altaWrxDb, permissionManager.Tenant.id);

                if (result != null)
                {
                    if (result.portRequestStatus == TelegenceNumbePortInStatus.CO.ToString() || result.portRequestStatus == TelegenceNumbePortInStatus.CF.ToString())
                    {
                        NumberPortInRepository NPIRR = new NumberPortInRepository(altaWrxDb);
                        NumberPortInRequest numberPortInRequest = new NumberPortInRequest();
                        numberPortInRequest.PortInRequestStatus = altaWrxDb.TelegencePortInRequestStatus.FirstOrDefault(x => x.Status.ToLower() == result.portRequestStatus.ToLower())?.id;
                        numberPortInRequest.ModifiedBy = user.Name;
                        numberPortInRequest.ModifiedDate = DateTime.UtcNow;
                        NPIRR.Object = numberPortInRequest;
                        NPIRR.Save(Session);

                        return new JsonResult { Data = "Success" };
                    }
                    else
                    {
                        return new JsonResult { Data = result.portRequestStatus };
                    }
                }
                else
                {
                    return new JsonResult { Data = "Error getting Request status. Please try again later." };
                }
            }
            catch (Exception e)
            {
                return new JsonResult { Data = e.Message };
            }
        }

        public async Task<ActionResult> TelegenceActivatePortInNumber(int numberPortInRequestId, string iccid, string imei, string ratePlanCode, int serviceProviderId = 0)
        {
            try
            {
                var portInRepository = new NumberPortInRepository(altaWrxDb);
                var numberPortInRequest = portInRepository.GetById(numberPortInRequestId);
                var subscriber = new RelatedParty();

                var portInDetails = JsonConvert.DeserializeObject<TelegencePortInRequest>(numberPortInRequest.RequestPayload);
                foreach (var relatedParty in portInDetails.RelatedParty)
                {
                    if (relatedParty.Role == "subscriber")
                    {
                        subscriber = relatedParty;
                    }
                }

                var tad = new TelegenceAdditionalDetails
                {
                    BillingAccountNo = portInDetails.BillingAccount.Id,
                    FirstName = subscriber.FirstName,
                    LastName = subscriber.LastName,
                    StreetNo = subscriber.Address.StreetNumber,
                    StreetName = subscriber.Address.StreetName,
                    StreetDirection = subscriber.Address.StreetDirection,
                    City = subscriber.Address.City,
                    State = subscriber.Address.State,
                    ZipCode = subscriber.Address.ZipCode,
                    IMEI = imei
                };

                var foundationAccList = altaWrxDb.usp_Telegence_Get_BillingAccounts();
                var telegenceDeviceRepository = new TelegenceDeviceRepository(altaWrxDb);
                var existingTelegenceDevice = telegenceDeviceRepository.GetByICCID(iccid);
                if (existingTelegenceDevice == null)
                {
                    var telegenceDevice = new TelegenceDevice
                    {
                        ServiceProviderId = serviceProviderId,
                        FoundationAccountNumber = foundationAccList.FirstOrDefault(x => x.BillingAccountNumber == tad.BillingAccountNo)?.FoundationAccountNumber,
                        BillingAccountNumber = tad.BillingAccountNo,
                        SubscriberNumber = numberPortInRequest.SubscriberNo,
                        ICCID = iccid,
                        IMEI = imei,
                        CreatedBy = user.Name,
                        CreatedDate = DateTime.UtcNow,
                        IsActive = true,
                        IsDeleted = false
                    };

                    telegenceDeviceRepository.SaveNew(Session, telegenceDevice);
                }

                return await UpdateStatus(iccid, TelegenceDeviceStatus.A.ToString(), ratePlanCode, tad.ZipCode, tad);
            }
            catch (Exception e)
            {
                return new JsonResult { Data = $"Exception: {e.Message}" };
            }
        }

        [HttpPost]
        public async Task<ActionResult> UpdateStatus(string iccid, string status, string ratePlanCode = null, string mdnZipCode = null, TelegenceAdditionalDetails tad = null, eBondingDeviceUpdateRequest eBondingDeviceUpdateRequest = null)
        {
            try
            {
                var device = altaWrxDb.MobilityDevices.Include("ServiceProvider").FirstOrDefault(x => x.ICCID == iccid);

                if (device == null)
                {
                    return ErrorMessage($"Could not find device with ICCID {iccid}");
                }

                var deviceTenant = altaWrxDb.MobilityDevice_Tenant.FirstOrDefault(x => x.MobilityDeviceId == device.id && x.TenantId == permissionManager.Tenant.id);

                var integrationType = (IntegrationType)device.ServiceProvider.IntegrationId;
                var serviceProviderId = device.ServiceProviderId;
                var devStatusFileRepo = new DeviceStatusUploadedFileRepository(altaWrxDb, permissionManager);
                var deviceStatusFile = devStatusFileRepo.Create(Session, DeviceStatusUploadedFileRepository.ADHOC_UPLOAD, serviceProviderId, deviceTenant?.SiteId);

                var allowedStatusList = altaWrxDb.DeviceStatus
                    .Where(x => !x.IsDeleted && x.IsActive && x.AllowsApiUpdate && x.IntegrationId == (int)integrationType)
                    .ToList()
                    .Select(x => x.Status.ToLower())
                    .ToList();

                var uploadedFileId = deviceStatusFile.id;
                var targetStatus = status;
                var additionalDetailsJSON = string.Empty;

                switch (integrationType)
                {
                    case IntegrationType.eBonding:
                        targetStatus = eBondingHelper.GetEbondingTargetStatus(status);
                        var careOrderRequest = eBondingHelper.GetEbondingCareOrderRequest(eBondingDeviceUpdateRequest, device, user);
                        additionalDetailsJSON = JsonConvert.SerializeObject(careOrderRequest, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
                        break;
                    case IntegrationType.Telegence:
                        additionalDetailsJSON = tad != null ? JsonConvert.SerializeObject(tad) : string.Empty;
                        break;
                }

                try
                {
                    var fileDetailRepo = new DeviceStatusUploadedFileDetailRepository(altaWrxDb, permissionManager);
                    fileDetailRepo.Create(SessionHelper.GetAuditByName(Session), integrationType, uploadedFileId, iccid, status, targetStatus, ratePlanCode, mdnZipCode, allowedStatusList, additionalDetailsJSON);
                    fileDetailRepo.Save();
                }
                catch (DbUpdateException ex)
                {
                    return ErrorMessage($"Could not queue the status change. {ex.Message}");
                }

                var model = new DeviceStatusUploadedFileModel();
                var validRecordExists = model.PrePostStatusCheck(altaWrxDb, permissionManager, Session, uploadedFileId);
                if (validRecordExists)
                {
                    var customObjectDbList = GetTenantCustomFields();
                    var awsAccessKey = AwsAccessKeyFromCustomObjects(customObjectDbList);
                    var awsSecretAccessKey = AwsSecretAccessKeyFromCustomObjects(customObjectDbList);
                    var deviceStatusChangeQueueName = StatusUploadQueueFromCustomObjects(customObjectDbList);

                    // validate custom fields
                    if (string.IsNullOrWhiteSpace(awsAccessKey) || string.IsNullOrWhiteSpace(awsSecretAccessKey) || string.IsNullOrWhiteSpace(deviceStatusChangeQueueName))
                    {
                        return ErrorMessage("Unable to complete. AWS setup is not complete for this Partner.");
                    }

                    var sqsHelper = new SqsHelper(awsAccessKey, awsSecretAccessKey);
                    var errorMessage = await sqsHelper.EnqueueDeviceStatusUpdateAsync(deviceStatusChangeQueueName, uploadedFileId, user.Username);

                    if (!string.IsNullOrWhiteSpace(errorMessage))
                    {
                        return ErrorMessage(errorMessage);
                    }

                    SessionHelper.SetAlert(Session, "Successfully started device status upload to Provider.");
                    SessionHelper.SetAlertType(Session, "success");
                }

                // Update JasperDeviceStatus_UploadedFile to Processed. 
                model.MarkFileAsProcessed(altaWrxDb, permissionManager, Session, uploadedFileId);

                SessionHelper.SetAlert(Session, "Successfully started device status upload.");
                SessionHelper.SetAlertType(Session, "success");
            }
            catch (Exception ex)
            {
                return ErrorMessage($"Error queuing the status change. {ex.Message}");
            }

            return Content("OK");
        }

        [HttpPost]
        public async Task<ActionResult> UpdateEbondingDevice(eBondingDeviceUpdateRequest deviceUpdateRequest)
        {
            try
            {
                var device = altaWrxDb.MobilityDevices.FirstOrDefault(x => x.ICCID == deviceUpdateRequest.ICCID);
                if (device == null)
                {
                    return ErrorMessage($"Could not find device with ICCID {deviceUpdateRequest.ICCID}");
                }

                if (!eBondingHelper.IsStatusUpdate(deviceUpdateRequest.ChangeRequestType))
                {
                    return await UpdateMobilityConfiguration(deviceUpdateRequest, device);
                }

                var result = await UpdateStatus(deviceUpdateRequest.ICCID, deviceUpdateRequest.ChangeRequestType, null, null, null, deviceUpdateRequest);
                SessionHelper.SetAlert(Session, "Successfully started device status upload.");
                SessionHelper.SetAlertType(Session, "success");
                return result;
            }
            catch (Exception ex)
            {
                return ErrorMessage($"Error queuing the status change. {ex.Message}");
            }
        }

        private async Task<ActionResult> UpdateMobilityConfiguration(eBondingDeviceUpdateRequest deviceUpdateRequest, IDevice device)
        {
            var careOrderRequest = eBondingHelper.GetEbondingCareOrderRequest(deviceUpdateRequest, device, user);
            var details = JsonConvert.SerializeObject(careOrderRequest, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });

            var change = CreateMobilityConfigurationChangeQueue(device.id, device.ServiceProviderId, details);

            var repository = new MobilityConfigurationChangeQueueRepository(altaWrxDb);
            var id = repository.SaveNew(change);
            var errorMessage = await EnqueueMobilityConfigurationChangeAsync(id);
            if (!string.IsNullOrEmpty(errorMessage))
            {
                return ErrorMessage(errorMessage);
            }
            int? tenantId = permissionManager.Tenant.id;
            OptimizationApiController optimizationApiController = new OptimizationApiController();
            optimizationApiController.SendTriggerAmopSync("mobility_inventory_live_sync", tenantId, null);
            SessionHelper.SetAlert(Session, "Successfully started device status upload.");
            SessionHelper.SetAlertType(Session, "success");
            return new HttpStatusCodeResult(HttpStatusCode.OK);
        }

        [HttpPost]
        public ActionResult UpdateMobilityCustomerRatePlan(int deviceId, decimal? customerDataAllocationMB, int? customerRatePlanId, int? customerRatePoolId, DateTime? effectiveDate)
        {
            try
            {
                int? tenantId = permissionManager.Tenant.id;
                // if effective date is today or empty, wait for sync daily to check and assign customer rate plan for device
                if (effectiveDate == null || effectiveDate < DateTime.Now)
                {
                    var deviceRepository = new MobilityDeviceRepository(altaWrxDb);
                    var device = deviceRepository.GetMobilityDeviceById(deviceId);
                    if (device == null)
                        return new JsonResult { Data = new { Success = false, Message = "Error updating Customer Rate Plan for device. Device not found." } };

                    var deviceTenant = altaWrxDb.MobilityDevice_Tenant.FirstOrDefault(x => x.MobilityDeviceId == device.id && x.TenantId == permissionManager.Tenant.id);
                    var customerRatePlan = altaWrxDb.JasperCustomerRatePlans.FirstOrDefault(x => x.id == customerRatePlanId && x.TenantId == permissionManager.Tenant.id && x.IsActive && !x.IsDeleted);
                    var customerRatePool = altaWrxDb.CustomerRatePools.FirstOrDefault(x => x.id == customerRatePoolId && x.TenantId == permissionManager.Tenant.id && x.IsActive && !x.IsDeleted);
                    var deviceActionHistories = new List<DeviceActionHistory>();
                    if (deviceTenant == null)
                    {
                        deviceTenant = new MobilityDevice_Tenant()
                        {
                            MobilityDeviceId = device.id,
                            TenantId = permissionManager.Tenant.id,
                            IsActive = true,
                            IsDeleted = false,
                            CreatedBy = SessionHelper.GetAuditByName(Session),
                            CreatedDate = DateTime.UtcNow
                        };

                        altaWrxDb.MobilityDevice_Tenant.Add(deviceTenant);
                        altaWrxDb.SaveChanges();
                    }
                    var previousCustomerRatePlan = deviceTenant.JasperCustomerRatePlan?.RatePlanName;
                    var previousCustomerRatePool = deviceTenant.CustomerRatePool?.Name;
                    if (customerRatePlanId != null && customerRatePlanId > 0)
                    {
                        deviceTenant.CustomerRatePlanId = customerRatePlanId;
                        deviceTenant.CustomerDataAllocationMB = customerDataAllocationMB;
                    }
                    else if (customerRatePlanId != CommonConstants.NO_CHANGE)
                    {
                        deviceTenant.CustomerRatePlanId = null;
                        deviceTenant.CustomerDataAllocationMB = null;
                    }

                    if (customerRatePoolId != null && customerRatePoolId >= 0)
                    {
                        deviceTenant.CustomerRatePoolId = customerRatePoolId;
                    }
                    else if (customerRatePoolId != CommonConstants.NO_CHANGE)
                    {
                        deviceTenant.CustomerRatePoolId = null;
                    }
                    deviceTenant.ModifiedDate = DateTime.UtcNow;
                    deviceTenant.ModifiedBy = SessionHelper.GetAuditByName(Session);

                    try
                    {
                        altaWrxDb.Entry(deviceTenant).State = EntityState.Modified;
                        altaWrxDb.SaveChanges();
                        if (previousCustomerRatePlan != customerRatePlan?.RatePlanName && customerRatePlanId != CommonConstants.NO_CHANGE)
                        {
                            var deviceActionHistory = new DeviceActionHistory()
                            {
                                ServiceProviderId = device.ServiceProviderId,
                                MobilityDeviceId = device.id,
                                ICCID = device.ICCID,
                                MSISDN = device.MSISDN,
                                PreviousValue = previousCustomerRatePlan,
                                CurrentValue = customerRatePlan?.RatePlanName,
                                ChangedField = CommonStrings.CustomerRatePlan,
                                ChangeEventType = CommonStrings.UpdateCustomerRatePlan,
                                DateOfChange = DateTime.UtcNow,
                                ChangedBy = SessionHelper.GetAuditByName(Session),
                                Username = device.Username,
                                CustomerAccountName = deviceTenant.Site.Name,
                                CustomerAccountNumber = deviceTenant.AccountNumber,
                                TenantId = permissionManager.Tenant.id,
                                IsActive = true,
                                IsDeleted = false
                            };
                            deviceActionHistories.Add(deviceActionHistory);
                        }
                        if (previousCustomerRatePool != customerRatePool?.Name && customerRatePoolId != CommonConstants.NO_CHANGE)
                        {
                            var deviceActionHistory = new DeviceActionHistory()
                            {
                                ServiceProviderId = device.ServiceProviderId,
                                MobilityDeviceId = device.id,
                                ICCID = device.ICCID,
                                MSISDN = device.MSISDN,
                                PreviousValue = previousCustomerRatePool,
                                CurrentValue = customerRatePool?.Name,
                                ChangedField = CommonStrings.CustomerRatePool,
                                ChangeEventType = CommonStrings.UpdateCustomerRatePlan,
                                DateOfChange = DateTime.UtcNow,
                                ChangedBy = SessionHelper.GetAuditByName(Session),
                                Username = device.Username,
                                CustomerAccountName = deviceTenant.Site.Name,
                                CustomerAccountNumber = deviceTenant.AccountNumber,
                                TenantId = permissionManager.Tenant.id,
                                IsActive = true,
                                IsDeleted = false
                            };
                            deviceActionHistories.Add(deviceActionHistory);
                        }
                        altaWrxDb.DeviceActionHistories.AddRange(deviceActionHistories);
                        altaWrxDb.SaveChanges();
                        altaWrxDb.usp_UpdateCrossProviderDeviceHistory(string.Empty, deviceTenant.MobilityDeviceId.ToString(), (int)PortalTypeEnum.Mobility, deviceTenant.TenantId, device.ServiceProviderId, effectiveDate);
                        OptimizationApiController optimizationApiController = new OptimizationApiController();
                        optimizationApiController.SendTriggerAmopSync("mobility_inventory_live_sync", tenantId, null);
                    }
                    catch (Exception ex)
                    {
                        return new JsonResult
                        {
                            Data = new
                            {
                                Success = false,
                                Message = "Error updating Customer Rate Plan and/or Customer Rate Pool for device"
                            }
                        };
                    }

                    var isMobilityHistorySuccessfulUpdate = UpdateMobilityDeviceHistory(device, deviceTenant, effectiveDate);
                    if (!isMobilityHistorySuccessfulUpdate)
                    {
                        return new JsonResult
                        {
                            Data = new
                            {
                                Success = false,
                                Message = "Successfully updated Customer Rate Plan and/or Customer Rate Pool but there was an error updating Mobility Device History."
                            }
                        };
                    }
                }
                else // save Device_CustomerRatePlanOrRatePool_Queue 
                {

                    var queue = new Device_CustomerRatePlanOrRatePool_Queue()
                    {
                        DeviceId = deviceId,
                        EffectiveDate = effectiveDate,
                        PortalType = (int)PortalTypes.Mobility,
                        CreatedBy = SessionHelper.GetAuditByName(Session),
                        CreatedDate = DateTime.UtcNow,
                        IsActive = true,
                        TenantId = tenantId
                    };

                    if (customerRatePlanId != null && (customerRatePlanId > 0 || customerRatePlanId == CommonConstants.NO_CHANGE))
                    {
                        queue.CustomerRatePlanId = customerRatePlanId;
                        queue.CustomerDataAllocationMB = customerDataAllocationMB;
                    }
                    if (customerRatePoolId != null && (customerRatePoolId > 0 || customerRatePoolId == CommonConstants.NO_CHANGE))
                    {
                        queue.CustomerRatePoolId = customerRatePoolId;
                    }
                    altaWrxDb.Device_CustomerRatePlanOrRatePool_Queue.Add(queue);
                    altaWrxDb.SaveChanges();
                    OptimizationApiController optimizationApiController = new OptimizationApiController();
                    optimizationApiController.SendTriggerAmopSync("mobility_inventory_live_sync", tenantId, null);
                }

                return new JsonResult
                {
                    Data = new
                    {
                        Success = true,
                        Message = "Successfully updated Customer Rate Plan and/or Customer Rate Pool."
                    }
                };
            }
            catch (Exception ex)
            {
                return new JsonResult
                {
                    Data = new
                    {
                        Success = false,
                        Message = $"Error updating Customer Rate Plan and/or Customer Rate Pool for device. {ex.Message}"
                    }
                };
            }
        }

        public ActionResult RevenueAssurance(string filter, Utils.CustomerAssignedFilter customerAssignedFilter = Utils.CustomerAssignedFilter.All,
            bool showVarianceOnly = true)
        {
            ViewBag.PageTitle = "Revenue Assurance";
            var tenantId = permissionManager.PermissionFilter.LoggedInTenantId;
            var mobilityRevenueAssuranceGroupModel = new MobilityRevenueAssuranceGroupModel();
            var mobilityRevenueAssuranceGroupRepository = new MobilityRevenueAssuranceGroupRepository(altaWrxDb, permissionManager);
            var group = mobilityRevenueAssuranceGroupRepository.GetMobilityRevenueAssuraceGroup(tenantId, filter, customerAssignedFilter, showVarianceOnly);
            if (group != null)
            {
                mobilityRevenueAssuranceGroupModel.Filter = filter;
                mobilityRevenueAssuranceGroupModel.CustomerAssignedFilter = customerAssignedFilter;
                mobilityRevenueAssuranceGroupModel.MobilityRevenueAssuranceGroup = group;
                mobilityRevenueAssuranceGroupModel.ShowVarianceOnly = showVarianceOnly;
            }
            return View(mobilityRevenueAssuranceGroupModel);
        }

        public ActionResult GetRevenueAssuranceByCustomer(string customerId, int page = 1, int pageSize = 25,
            string sort = null, string sortDir = null, bool showVarianceOnly = true)
        {
            var tenantId = permissionManager.PermissionFilter.LoggedInTenantId;
            var mobilityRevServiceProductRepository = new MobilityRevServiceProductRepository(altaWrxDb);
            var totalDevicesCount = mobilityRevServiceProductRepository.GetMobilityRevServiceProductByCustomerCount(tenantId, customerId, showVarianceOnly);
            var DIDsFilter = new List<string>();
            if (!customerId.Equals(CommonConstants.UNASSIGNED))
            {
                DIDsFilter = ListHelper.GetIdentifierListFromDIDs(altaWrxDb, customerId, permissionManager.Tenant.id, base64Service, IsProduction, totalDevicesCount);
            }
            var revServiceProductsByCustomer = mobilityRevServiceProductRepository.GetMobilityRevServiceProductByCustomer(tenantId, customerId, showVarianceOnly, page, sort, sortDir, DIDsFilter: DIDsFilter);
            var revServiceProducts = PagedList.ToPagedList(revServiceProductsByCustomer, totalDevicesCount);
            var activeDevicesCount = mobilityRevServiceProductRepository.GetActiveRevServiceProductByCustomerCount(tenantId, customerId, showVarianceOnly);
            var isVariant = revServiceProductsByCustomer.Any(rsp => rsp.IsActiveStatus != rsp.RevIsActiveStatus);
            var mobilityRevenueAssuranceDeviceListModel = new MobilityRevenueAssuranceDeviceListModel
            {
                RevCustomerId = customerId,
                RevServiceProducts = revServiceProducts,
                ActiveDevicesCount = activeDevicesCount,
                TotalDevicesCount = totalDevicesCount,
                IsVariant = isVariant
            };
            return PartialView("_RevCustomerMobilityRevenueAssurance", mobilityRevenueAssuranceDeviceListModel);
        }

        public ActionResult GetRevenueAssuranceByProductId(string productIds, string customerId, bool showVarianceOnly)
        {
            var tenantId = permissionManager.PermissionFilter.LoggedInTenantId;
            var productIdList = productIds.Split(',');
            var mobilityRevServiceProductRepository = new MobilityRevServiceProductRepository(altaWrxDb);
            var revServiceProductsByCustomers =
                mobilityRevServiceProductRepository.GetMobilityRevDisconnectServiceProductByCustomer(tenantId, customerId, showVarianceOnly, productIdList, true).Select(x => x.ToRevDisconnectSDDevice()).ToList();

            var dataResult = new List<RevSelectedServiceProductWithDisconnectDate>();
            var dsHistoryRepository = new DeviceStatusHistoryRepository(altaWrxDb, tenantId, PORTAL_TYPE);

            foreach (var revSP in revServiceProductsByCustomers)
            {
                var dsHistory = dsHistoryRepository.GetMobilityDeviceByICCID(revSP.ICCID);
                var spSelected = new RevSelectedServiceProductWithDisconnectDate()
                {
                    ServiceProduct = revSP,
                    DisconnectCarrierDate = dsHistory?.DateOfChange.Value,
                };
                dataResult.Add(spSelected);
            }

            var model = new RevDisconnectServiceProductSelectedModel()
            {
                records = dataResult,
                RevCustomerId = customerId,
                ShowVarianceOnly = showVarianceOnly
            };

            return PartialView("_RevDisconnectServiceProductSelected", model);
        }

        public FileContentResult RevenueAssuranceExport(string filter,
            Utils.CustomerAssignedFilter customerAssignedFilter = Utils.CustomerAssignedFilter.All, bool showVarianceOnly = true)
        {
            if (permissionManager.UserCannotAccess(Session, ModuleEnum.Mobility))
                return null;

            var tenantId = permissionManager.PermissionFilter.LoggedInTenantId;
            var mobilityRevServiceProductRepository = new MobilityRevServiceProductRepository(altaWrxDb);
            var revServiceProducts = mobilityRevServiceProductRepository.GetMobilityRevServiceProductExport(tenantId, filter, customerAssignedFilter, showVarianceOnly)
                .Select(rs => rs.ToRevenueAssuranceExport()).ToList();

            var data = revServiceProducts.ToDataSet("Mobility Revenue Assurance");
            data.Tables["Mobility Revenue Assurance"].Columns.Remove("Communication Plan");
            var bytes = ExcelUtilities.Export(data);

            return File(bytes, ExcelContentType, $"Mobility_Revenue_Assurance_{FileNameTimestamp()}.{ExcelFileExtension}");
        }

        [HttpPost]
        public async Task<ActionResult> AssociateCustomer(BulkChangeAssociateCustomerModel model)
        {
            if (!permissionManager.UserCanCreate(Session, ModuleEnum.Mobility))
            {
                return new HttpStatusCodeResult(HttpStatusCode.Forbidden);
            }

            try
            {
                var bulkchangeId = await BuildAssociateCustomer(model);

                await ProcessBulkChange(bulkchangeId);
                return new JsonResult { Data = new { Success = true, ChangeId = bulkchangeId } };
            }
            catch (Exception e)
            {
                return new JsonResult { Data = new { Success = false, e.Message } };
            }
        }

        private async Task<long> BuildAssociateCustomer(BulkChangeAssociateCustomerModel model)
        {
            model.Devices = model.Devices.Distinct().ToArray();
            var changeRepository = new DeviceBulkChangeRepository(altaWrxDb, permissionManager);
            var changeType = DeviceChangeType.CustomerAssignment;
            var useCarrierActivation = model.EffectiveDate == null ? true : false;

            var bulkChange = new DeviceBulkChange
            {
                ChangeRequestTypeId = (int)changeType,
                ServiceProviderId = model.ServiceProviderId,
                TenantId = permissionManager.PermissionFilter.LoggedInTenantId,
                Status = BulkChangeStatus.NEW,
                CreatedDate = DateTime.UtcNow,
                CreatedBy = SessionHelper.GetAuditByName(Session),
                IsActive = true,
                IsDeleted = false,
                Mobility_DeviceChange =
                    BuildAssociateCustomerDeviceChanges(altaWrxDb, Session, permissionManager, model, useCarrierActivation).ToList()
            };
            var bulkChangeId = changeRepository.CreateBulkChange(bulkChange);

            if (!model.CreateRevService && !model.AddCarrierRatePlan)
            {
                await ProcessBulkAssociateAMOP(bulkChange.id, model);
            }

            return bulkChange.id;
        }

        public async Task<ActionResult> AssociateAmopCustomer(BulkChangeAssociateNonRevCustomerModel model)
        {
            if (!permissionManager.UserCanCreate(Session, ModuleEnum.M2M) && permissionManager.UserCanAccess(Session, ModuleEnum.RevCustomers))
            {
                return new HttpStatusCodeResult(HttpStatusCode.Forbidden);
            }
            try
            {
                model.Devices = model.Devices.Distinct().ToArray();
                var changeRepository = new DeviceBulkChangeRepository(altaWrxDb, permissionManager);
                var changeType = DeviceChangeType.CustomerAssignment;
                var changes = BuildAssociateAmopCustomerDeviceChanges(altaWrxDb, Session, permissionManager, model).ToList();
                var bulkChange = new DeviceBulkChange
                {
                    ChangeRequestTypeId = (int)changeType,
                    ServiceProviderId = model.ServiceProviderId,
                    TenantId = permissionManager.PermissionFilter.LoggedInTenantId,
                    Status = changes.Any(change => !change.IsProcessed) ? BulkChangeStatus.NEW : BulkChangeStatus.PROCESSED,
                    CreatedDate = DateTime.UtcNow,
                    CreatedBy = SessionHelper.GetAuditByName(Session),
                    IsActive = true,
                    IsDeleted = false,
                    Mobility_DeviceChange = changes
                };
                changeRepository.CreateBulkChange(bulkChange);
                await ProcessBulkAssociateAMOPCustomer(bulkChange.id);
                return new JsonResult { Data = new { Success = true } };
            }
            catch (Exception e)
            {
                return new JsonResult { Data = new { Success = false, e.Message } };
            }
        }
        private async Task ProcessBulkAssociateAMOPCustomer(long bulkChangeId)
        {
            var mobilityDeviceBulkChangeRepository = new MobilityDeviceChangeRepository(altaWrxDb, permissionManager);
            var changes = mobilityDeviceBulkChangeRepository.GetUnprocessedChanges(bulkChangeId);
            {
                var change = changes.First();
                var dataTableUpdates = BuildTableAssignNonRevCustomer(changes);
                await DeviceBulkChangeAssignNonRevCustomer(permissionManager.AltaworxCentralConnectionStringWithoutEF, dataTableUpdates, bulkChangeId, change.id);
            }
        }
        private DataTable BuildTableAssignNonRevCustomer(IList<Mobility_DeviceChange> changes)
        {
            var table = new DataTable();
            table.Columns.Add("DeviceId");
            table.Columns.Add("TenantId");
            table.Columns.Add("SiteId", typeof(int));
            foreach (var change in changes)
            {
                if (!string.IsNullOrWhiteSpace(change.DeviceChangeRequest?.ChangeRequest))
                {
                    var associateNonRevCustomerModel = JsonConvert.DeserializeObject<BulkChangeAssociateNonRevCustomerModel>(change.DeviceChangeRequest?.ChangeRequest);
                    var dataRow = AddDataToTableAssignNonRev(table, change, associateNonRevCustomerModel);
                    table.Rows.Add(dataRow);
                }
            }
            return table;
        }
        private DataRow AddDataToTableAssignNonRev(DataTable table, Mobility_DeviceChange detailRecord, BulkChangeAssociateNonRevCustomerModel nonRevCustomerModel)
        {
            var dr = table.NewRow();
            dr[0] = detailRecord.DeviceId;
            dr[1] = nonRevCustomerModel.TenantId;
            dr[2] = nonRevCustomerModel.SiteId;
            return dr;
        }
        private async Task DeviceBulkChangeAssignNonRevCustomer(string CentralDbConnectionString, DataTable table, long bulkChangeId, long deviceChangeId)
        {
            DeviceChangeResult<string, string> dbResult;
            var deviceBulkChangeLogRepository = new DeviceBulkChangeLogRepository(altaWrxDb);
            try
            {
                using (var conn = new SqlConnection(CentralDbConnectionString))
                {
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandText = "dbo.usp_DeviceBulkChange_Assign_Non_Rev_Customer";
                        SqlParameter newrecordParam = cmd.Parameters.Add("@UpdatedValues", SqlDbType.Structured);
                        cmd.Parameters["@UpdatedValues"].Value = table;
                        cmd.Parameters["@UpdatedValues"].TypeName = "dbo.UpdateM2MDeviceMobilityDeviceSiteType";
                        cmd.Parameters.AddWithValue("@bulkChangeId", bulkChangeId);
                        cmd.Parameters.AddWithValue("@portalTypeId", PortalTypes.Mobility);
                        conn.Open();

                        await cmd.ExecuteNonQueryAsync();
                    }
                }
                dbResult = new DeviceChangeResult<string, string>()
                {
                    ActionText = "usp_DeviceBulkChange_Assign_Non_Rev_Customer",
                    HasErrors = false,
                    RequestObject = $"bulkChangeId: {bulkChangeId}",
                    ResponseObject = "OK"
                };
            }
            catch (Exception ex)
            {
                Log.Error($"Error Executing Stored Procedure usp_DeviceBulkChange_Assign_Non_Rev_Customer: {ex.Message} {ex.StackTrace}");
                var logId = Guid.NewGuid();
                dbResult = new DeviceChangeResult<string, string>()
                {
                    ActionText = "usp_DeviceBulkChange_Assign_Non_Rev_Customer",
                    HasErrors = true,
                    RequestObject = $"bulkChangeId: {bulkChangeId}",
                    ResponseObject = $"Error Executing Stored Procedure. Ref: {logId}"
                };
            }
            deviceBulkChangeLogRepository.AddMobilityLogEntry(new CreateMobilityDeviceBulkChangeLog()
            {
                BulkChangeId = bulkChangeId,
                ErrorText = dbResult.HasErrors ? dbResult.ResponseObject : null,
                HasErrors = dbResult.HasErrors,
                LogEntryDescription = "AssignNonRevCustomer: Update AMOP",
                MobilityDeviceChangeId = deviceChangeId,
                ProcessBy = "AltaworxDeviceBulkChange",
                ProcessedDate = DateTime.UtcNow,
                ResponseStatus = dbResult.HasErrors ? BulkChangeStatus.ERROR : BulkChangeStatus.PROCESSED,
                RequestText = dbResult.ActionText + Environment.NewLine + dbResult.RequestObject,
                ResponseText = dbResult.ResponseObject
            });

            var bulkChange = altaWrxDb.DeviceBulkChanges.Find(bulkChangeId);
            var processedBy = SessionHelper.GetAuditByName(Session);
            var processedDate = DateTime.UtcNow;
            bulkChange.Status = BulkChangeStatus.PROCESSED;
            bulkChange.ProcessedBy = processedBy;
            bulkChange.ProcessedDate = processedDate;
            bulkChange.ModifiedBy = processedBy;
            bulkChange.ModifiedDate = processedDate;
            await altaWrxDb.SaveChangesAsync();
        }
        internal static IEnumerable<Mobility_DeviceChange> BuildAssociateAmopCustomerDeviceChanges(AltaWorxCentral_Entities awxDb,
        HttpSessionStateBase session, PermissionManager permissionManager, BulkChangeAssociateNonRevCustomerModel model)
        {
            var createdBy = SessionHelper.GetAuditByName(session);
            var devicesByPhoneNumbers = GetDevicesByNumber(awxDb, model.ServiceProviderId, model.Devices);
            var archivedMSISDNs = CheckDevicesArchived(awxDb, model.ServiceProviderId, model.Devices);

            var deviceChanges = new List<Mobility_DeviceChange>();
            var changeRequest = CreateAssociateAmopCustomerChangeRequest(model, permissionManager.Tenant.id);

            foreach (var phoneNumber in model.Devices)
            {
                if (!string.IsNullOrWhiteSpace(phoneNumber))
                {
                    if (!devicesByPhoneNumbers.ContainsKey(phoneNumber) || !devicesByPhoneNumbers.TryGetValue(phoneNumber, out var device))
                    {
                        if (archivedMSISDNs.Contains(phoneNumber))
                        {
                            deviceChanges.Add(CreateDeviceChangeError(phoneNumber, string.Format(CommonStrings.MobilityDeviceSubscriberNumberIsArchivedError, phoneNumber), createdBy));
                        }
                        else
                        {
                            deviceChanges.Add(CreateDeviceChangeError(phoneNumber, string.Format(CommonStrings.MobilityDeviceSubscriberNumberNotExistError, phoneNumber), createdBy));
                        }
                    }
                    else
                    {
                        deviceChanges.Add(new Mobility_DeviceChange(changeRequest, device.id, device.ICCID, phoneNumber, createdBy));
                    }
                }
            }

            return deviceChanges;
        }
        private static string CreateAssociateAmopCustomerChangeRequest(BulkChangeAssociateNonRevCustomerModel model, int tenantId)
        {
            return JsonConvert.SerializeObject(new BulkChangeAssociateNonRevCustomerModel()
            {
                Devices = model.Devices,
                ServiceProviderId = model.ServiceProviderId,
                TenantId = tenantId,
                SiteId = model.SiteId
            });
        }


        public FileContentResult MobilityBulkChangeExport(int bulkChangeId)
        {
            if (!permissionManager.UserCanAccessPortalTypeModule(Session, PORTAL_TYPE))
            {
                return null;
            }
            var changeRepository = new DeviceBulkChangeRepository(altaWrxDb, permissionManager);
            var bulkChange = changeRepository.GetBulkChange(bulkChangeId);
            if (bulkChange != null && bulkChange.PortalTypeId == (int)PORTAL_TYPE)
            {
                var lineItems = altaWrxDb.usp_MobilityBulkChangeLogExport((int)bulkChange.id).ToList();
                var listTemp = new List<BulkChangeDetailExportModel>();
                foreach (var lineItem in lineItems)
                {
                    var itemAdd = new BulkChangeDetailExportModel
                    {
                        ChangeRequestType = lineItem.ChangeRequestType,
                        CreatedBy = lineItem.CreatedBy,
                        CreatedDate = lineItem.CreatedDate.ToString("MM/dd/yyyy"),
                        ProcessedBy = lineItem.ProcessedBy,
                        Status = lineItem.Status,
                        StatusDetails = lineItem.StatusDetails,
                        SubscriberNumber = lineItem.SubscriberNumber,
                        ProcessedDate = null
                    };

                    if (lineItem.ProcessedDate != null)
                    {
                        itemAdd.ProcessedDate = ((DateTime)lineItem.ProcessedDate).ToString("MM/dd/yyyy");
                    }
                    listTemp.Add(itemAdd);
                }
                var data = listTemp.ToDataSet("Bulk Change Detail");

                var bytes = ExcelUtilities.Export(data);

                return File(bytes, ExcelContentType, $"ReportMobilityInventory_{FileNameTimestamp()}.{ExcelFileExtension}");
            }
            else
            {
                return null;
            }
        }

        internal static IEnumerable<Mobility_DeviceChange> BuildAssociateCustomerDeviceChanges(AltaWorxCentral_Entities awxDb,
            HttpSessionStateBase session, PermissionManager permissionManager, BulkChangeAssociateCustomerModel model, bool useCarrierActivation = false)
        {
            var tenantTimeZone = TimeZoneHelper.GetTimeZoneInfo(permissionManager.AltaworxCentralConnectionString);

            var createdBy = SessionHelper.GetAuditByName(session);
            var devicesByPhoneNumbers = GetDevicesByNumber(awxDb, model.ServiceProviderId, model.Devices);
            var archivedMSISDNs = CheckDevicesArchived(awxDb, model.ServiceProviderId, model.Devices);

            var revCustomerRepository = new RevCustomerRepository(awxDb, permissionManager.Tenant.id);

            var deviceChanges = new List<Mobility_DeviceChange>();
            var revCustomer = revCustomerRepository.GetByRevCustomerId(model.RevCustomerId);

            if (revCustomer == null)
                return deviceChanges;

            var integrationAuthenticationId = revCustomer.IntegrationAuthenticationId;

            var revServiceRepository = new RevServiceRepository(awxDb, integrationAuthenticationId.GetValueOrDefault());
            var revServiceProductRepository = new RevServiceProductRepository(awxDb);
            var carrierActivationHelper = new CarrierActivationHelper(awxDb);
            foreach (var phoneNumber in model.Devices)
            {
                if (!string.IsNullOrWhiteSpace(phoneNumber))
                {
                    if (!devicesByPhoneNumbers.ContainsKey(phoneNumber) || !devicesByPhoneNumbers.TryGetValue(phoneNumber, out var device))
                    {
                        if (archivedMSISDNs.Contains(phoneNumber))
                        {
                            deviceChanges.Add(CreateDeviceChangeError(phoneNumber, string.Format(CommonStrings.MobilityDeviceSubscriberNumberIsArchivedError, phoneNumber), createdBy));
                        }
                        else
                        {
                            deviceChanges.Add(CreateDeviceChangeError(phoneNumber, string.Format(CommonStrings.MobilityDeviceSubscriberNumberNotExistError, phoneNumber), createdBy));
                        }
                    }
                    else if (model.CreateRevService && CheckRevServiceStatus(revServiceRepository, revServiceProductRepository, device.id, permissionManager.Tenant.id, revCustomer.RevCustomerId))
                    {
                        deviceChanges.Add(CreateDeviceChangeError(phoneNumber, string.Format(CommonStrings.ActiveServiceLineSubscriberNumberError, phoneNumber), createdBy));
                    }
                    else
                    {
                        if (useCarrierActivation)
                        {
                            var activatedDate = carrierActivationHelper.MobilityActivationDateFromNumber(device.MSISDN, tenantTimeZone);
                            if (activatedDate == null)
                            {
                                activatedDate = model.ActivatedDate;
                            }

                            model.ActivatedDate = activatedDate;
                            model.EffectiveDate = activatedDate;
                        }
                        deviceChanges.Add(new Mobility_DeviceChange(CreateAssociateCustomerChangeRequest(model, device, integrationAuthenticationId.GetValueOrDefault()), device.id, device.ICCID, phoneNumber, createdBy));
                    }
                }
            }

            return deviceChanges;
        }

        internal static IEnumerable<Mobility_DeviceChange> BuildUpdateICCIDorIMEI(AltaWorxCentral_Entities awxDb,
           HttpSessionStateBase session, PermissionManager permissionManager, BulkchangeUpdateICCIDorIMEI model)
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
                    if (!devicesByPhoneNumbers.ContainsKey(phoneNumber) || !devicesByPhoneNumbers.TryGetValue(phoneNumber, out var device))
                    {
                        if (archivedMSISDNs.Contains(phoneNumber))
                        {
                            deviceChanges.Add(CreateDeviceChangeError(phoneNumber, string.Format(CommonStrings.MobilityDeviceSubscriberNumberIsArchivedError, phoneNumber), createdBy));
                        }
                        else
                        {
                            deviceChanges.Add(CreateDeviceChangeError(phoneNumber, string.Format(CommonStrings.MobilityDeviceSubscriberNumberNotExistError, phoneNumber), createdBy));
                        }
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

                        deviceChanges.Add(new Mobility_DeviceChange(CreateUpdateICCIDorIMEIChangeRequest(newICCID, newIMEI, device, device.ServiceZipCode, CommonStrings.UpdateICCIDorIMEIReasonCode, device.TechnologyType, model), device.id, device.ICCID, phoneNumber, createdBy));
                    }
                }
            }
            return deviceChanges;
        }

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
            var changeEquipmentRequest = new TelegenceUpdateICCIDorIMEIRequest()
            {
                ServiceCharacteristic = characteristicList
            };
            var request = new BulkChangeStatusUpdateRequest<TelegenceUpdateICCIDorIMEIRequest>
            {
                Request = changeEquipmentRequest
            };
            if (model.IsChangeCustomerRatePlan)
            {
                var assignCustomer = new BulkChangeAssociateCustomer()
                {
                    AddCustomerRatePlan = model.IsChangeCustomerRatePlan,
                    CustomerRatePlan = model.CustomerRatePlanId.ToString(),
                    CustomerRatePool = model.CustomerRatePoolId.ToString()
                };
                request.RevService = assignCustomer;
            }
            return JsonConvert.SerializeObject(request);
        }

        internal static IEnumerable<Mobility_DeviceChange> BuildAssociateCustomerByUploadFile(AltaWorxCentral_Entities awxDb,
           HttpSessionStateBase session, PermissionManager permissionManager, BulkChangeAssociateCustomerModel model, bool useCarrierActivation = false)
        {
            var createdBy = SessionHelper.GetAuditByName(session);
            var devicesByIccids = GetDevicesByIccid(awxDb, model.ServiceProviderId, model.Devices);
            var archivedIccids = CheckDevicesArchivedByIccid(awxDb, model.ServiceProviderId, model.Devices);

            var revCustomerRepository = new RevCustomerRepository(awxDb, permissionManager.Tenant.id);

            var deviceChanges = new List<Mobility_DeviceChange>();
            var revCustomer = revCustomerRepository.GetByRevCustomerId(model.RevCustomerId);

            if (revCustomer == null)
                return deviceChanges;

            var integrationAuthenticationId = revCustomer.IntegrationAuthenticationId;

            var revServiceRepository = new RevServiceRepository(awxDb, integrationAuthenticationId.GetValueOrDefault());
            var revServiceProductRepository = new RevServiceProductRepository(awxDb);
            var carrierActivationHelper = new CarrierActivationHelper(awxDb);
            foreach (var iccid in model.Devices)
            {
                if (!string.IsNullOrEmpty(iccid))
                {
                    devicesByIccids.TryGetValue(iccid, out var device);
                    if (!devicesByIccids.ContainsKey(iccid))
                    {
                        if (archivedIccids.Contains(iccid))
                        {
                            deviceChanges.Add(CreateDeviceChangeErrorByIccid(iccid, string.Format(CommonStrings.MobilityDeviceICCIDIsArchivedError, iccid), createdBy));
                        }
                        else
                        {
                            deviceChanges.Add(CreateDeviceChangeErrorByIccid(iccid, string.Format(CommonStrings.MobilityDeviceICCIDNotExistError, iccid), createdBy));
                        }
                    }
                    else if (model.CreateRevService && CheckRevServiceStatus(revServiceRepository, revServiceProductRepository, device.id, permissionManager.Tenant.id, revCustomer.RevCustomerId))
                    {
                        deviceChanges.Add(CreateDeviceChangeErrorByIccid(iccid, string.Format(CommonStrings.ActiveServiceLineICCIDError, iccid), createdBy));
                    }
                    else
                    {
                        if (useCarrierActivation)
                        {
                            var activatedDate = carrierActivationHelper.MobilityActivationDateFromIccid(device.ICCID) ?? model.ActivatedDate;
                            model.ActivatedDate = activatedDate;
                            model.EffectiveDate = activatedDate;
                        }
                        deviceChanges.Add(new Mobility_DeviceChange(CreateAssociateCustomerChangeRequest(model, device, integrationAuthenticationId.GetValueOrDefault(), true), device.id, device.ICCID, device.MSISDN, createdBy));
                    }
                }
            }

            return deviceChanges;
        }

        internal static IEnumerable<Mobility_DeviceChange> BuildAssociateCustomerChangeAfterActivation(AltaWorxCentral_Entities awxDb,
           HttpSessionStateBase session, PermissionManager permissionManager, BulkChangeAssociateCustomerModel model, bool useCarrierActivation = false)
        {
            var createdBy = SessionHelper.GetAuditByName(session);
            var revCustomerRepository = new RevCustomerRepository(awxDb, permissionManager.Tenant.id);

            var deviceChanges = new List<Mobility_DeviceChange>();
            var revCustomer = revCustomerRepository.GetByRevCustomerId(model.RevCustomerId);

            if (revCustomer == null)
                return deviceChanges;

            var integrationAuthenticationId = revCustomer.IntegrationAuthenticationId;
            foreach (var iccid in model.Devices)
            {
                if (!string.IsNullOrWhiteSpace(iccid))
                {
                    deviceChanges.Add(new Mobility_DeviceChange(CreateAssociateCustomerActivateChangeRequest(model, iccid, integrationAuthenticationId.GetValueOrDefault(), true), null, iccid, string.Empty, createdBy));

                }
            }

            return deviceChanges;
        }

        private static bool CheckRevServiceStatus(RevServiceRepository revServiceRepository, RevServiceProductRepository revServiceProductRepository, int deviceId, int tenantId, string revCustomerId)
        {
            var revService = revServiceRepository.GetRevServiceByMobilityDeviceId(deviceId, tenantId);
            if (revService == null)
                return false;
            if (revService.DisconnectedDate == null)
            {
                //if a service does not have service product. Or the service product are DISCONNECTED => Create a new service line
                var serviceProducts = revServiceProductRepository.GetRevServiceProductByRevService(revService.RevServiceId);
                return serviceProducts.Any(x => x.Status != RevServiceProductStatus.DISCONNECTED.ToString() && x.CustomerId == revCustomerId);
            }
            return false;
        }

        private static string CreateAssociateCustomerChangeRequest(BulkChangeAssociateCustomerModel model, MobilityDevice device, int integrationAuthenticationId, bool uploadByFile = false)
        {
            if (model.ActivatedDate == null)
            {
                model.ActivatedDate = DateTime.UtcNow;
            }
            return JsonConvert.SerializeObject(new BulkChangeAssociateCustomer()
            {
                Number = uploadByFile ? String.Empty : device.MSISDN,
                ICCID = device.ICCID,
                RevCustomerId = model.RevCustomerId,
                DeviceId = device.id,
                CreateRevService = model.CreateRevService,
                ServiceTypeId = model.ServiceTypeId.GetValueOrDefault(0),
                RevPackageId = model.RevPackageId,
                RevProductIdList = model.RevProductIdList,
                RateList = model.RateList,
                Prorate = model.Prorate,
                Description = model.Description,
                EffectiveDate = model.EffectiveDate,
                AddCustomerRatePlan = model.AddCustomerRatePlan,
                CustomerRatePlan = model.CustomerRatePlan,
                CustomerRatePool = model.CustomerRatePool,
                IntegrationAuthenticationId = integrationAuthenticationId,
                ProviderId = model.ProviderId,
                ActivatedDate = model.ActivatedDate,
                UsagePlanGroupId = model.UsagePlanGroupId
            });
        }

        private static string CreateAssociateCustomerActivateChangeRequest(BulkChangeAssociateCustomerModel model, string iccid, int integrationAuthenticationId, bool uploadByFile = false)
        {
            if (model.ActivatedDate == null)
            {
                model.ActivatedDate = DateTime.UtcNow;
            }
            return JsonConvert.SerializeObject(new BulkChangeAssociateCustomer()
            {
                Number = string.Empty,
                ICCID = iccid,
                RevCustomerId = model.RevCustomerId,
                DeviceId = 0,
                CreateRevService = model.CreateRevService,
                ServiceTypeId = model.ServiceTypeId.GetValueOrDefault(0),
                RevProductIdList = model.RevProductIdList,
                RateList = model.RateList,
                Prorate = model.Prorate,
                Description = model.Description,
                EffectiveDate = model.EffectiveDate,
                AddCustomerRatePlan = model.AddCustomerRatePlan,
                CustomerRatePlan = model.CustomerRatePlan,
                CustomerRatePool = model.CustomerRatePool,
                IntegrationAuthenticationId = integrationAuthenticationId,
                ProviderId = model.ProviderId,
                ActivatedDate = model.ActivatedDate,
                UsagePlanGroupId = model.UsagePlanGroupId
            });
        }

        internal static async Task<object> ProcessBulkChange(AmopBaseController controller, UrlHelper url, AltaWorxCentral_Entities awxDb, HttpSessionStateBase session, long id)
        {
            var bulkChange = awxDb.DeviceBulkChanges.Find(id);
            if (bulkChange == null)
            {
                return new HttpNotFoundResult();
            }

            var customObjectDbList = controller.GetTenantCustomFields();
            var awsAccessKey = controller.AwsAccessKeyFromCustomObjects(customObjectDbList);
            var awsSecretAccessKey = controller.AwsSecretAccessKeyFromCustomObjects(customObjectDbList);
            var queueName = ValueFromCustomObjects(customObjectDbList, "Device Bulk Change Queue");
            var sqsHelper = new SqsHelper(awsAccessKey, awsSecretAccessKey);
            var errorMessage = await sqsHelper.EnqueueBulkChangeAsync(queueName, id);
            if (!string.IsNullOrWhiteSpace(errorMessage))
            {
                return new { errors = new[] { $"An error occurred: {errorMessage}" } };
            }

            var processedBy = SessionHelper.GetAuditByName(session);
            var processedDate = DateTime.UtcNow;
            bulkChange.Status = BulkChangeStatus.PROCESSING;
            bulkChange.ProcessedBy = processedBy;
            bulkChange.ProcessedDate = processedDate;
            bulkChange.ModifiedBy = processedBy;
            bulkChange.ModifiedDate = processedDate;
            await awxDb.SaveChangesAsync();

            return new { redirectUrl = url.Action("BulkChange", "Mobility", new { id }) };
        }

        private async Task ProcessBulkAssociateAMOP(long bulkChangeId, BulkChangeAssociateCustomerModel model)
        {
            var mobilityDeviceChangeRepository = new MobilityDeviceChangeRepository(altaWrxDb, permissionManager);
            var deviceChanges = mobilityDeviceChangeRepository.GetUnprocessedChanges(bulkChangeId);
            var sessionUser = SessionHelper.GetAuditByName(Session);
            var deviceIdList = new List<string>();

            var mobilityDeviceRepository = new MobilityDeviceRepository(altaWrxDb);
            var revCustomerRepository = new RevCustomerRepository(altaWrxDb, permissionManager.Tenant.id);
            var siteRepository = new SiteRepository(altaWrxDb, permissionManager);
            var revSites = siteRepository.GetAllRev(null, true);

            var deviceBulkChangeLogRepository = new DeviceBulkChangeLogRepository(altaWrxDb);

            foreach (var deviceChange in deviceChanges)
            {
                var changeRequest = JsonConvert.DeserializeObject<BulkChangeAssociateCustomerModel>(deviceChange.DeviceChangeRequest?.ChangeRequest);
                var changeRequestDateIsInvalid = (changeRequest.EffectiveDate == null || (changeRequest.EffectiveDate >= DateTime.MinValue && changeRequest.EffectiveDate?.ToUniversalTime() <= DateTime.UtcNow));
                var mobilityDevice = mobilityDeviceRepository.GetMobilityDeviceById(deviceChange.DeviceId.GetValueOrDefault());

                if (mobilityDevice == null)
                {
                    string statusMessage = "AMOP Device not found";

                    deviceChange.Status = BulkChangeStatus.ERROR;
                    deviceChange.HasErrors = true;
                    deviceChange.StatusDetails = statusMessage;
                    deviceChange.ModifiedDate = DateTime.UtcNow;
                    deviceChange.ModifiedBy = sessionUser;

                    // change detail log
                    deviceBulkChangeLogRepository.AddMobilityLogEntry(new Amop.Core.Models.DeviceBulkChange.CreateMobilityDeviceBulkChangeLog()
                    {
                        BulkChangeId = bulkChangeId,
                        HasErrors = true,
                        LogEntryDescription = "AMOP Customer Assignment",
                        MobilityDeviceChangeId = deviceChange.id,
                        ProcessBy = sessionUser,
                        ProcessedDate = DateTime.UtcNow,
                        RequestText = deviceChange.DeviceChangeRequest?.ChangeRequest,
                        ResponseStatus = BulkChangeStatus.ERROR,
                        ErrorText = statusMessage,
                        ResponseText = $"{CommonStrings.PreflightCheckError} {statusMessage}"
                    });

                    continue;
                }

                var revCustomer = revCustomerRepository.GetByRevCustomerId(changeRequest.RevCustomerId);
                var site = revSites.FirstOrDefault(x => x.RevCustomerId == revCustomer.id);

                if (site == null)
                {
                    string statusMessage = "Rev Customer not found";

                    deviceChange.Status = BulkChangeStatus.ERROR;
                    deviceChange.HasErrors = true;
                    deviceChange.StatusDetails = statusMessage;
                    deviceChange.ModifiedDate = DateTime.UtcNow;
                    deviceChange.ModifiedBy = sessionUser;

                    // change detail log
                    deviceBulkChangeLogRepository.AddMobilityLogEntry(new Amop.Core.Models.DeviceBulkChange.CreateMobilityDeviceBulkChangeLog()
                    {
                        BulkChangeId = bulkChangeId,
                        HasErrors = true,
                        LogEntryDescription = "AMOP Customer Assignment",
                        MobilityDeviceChangeId = deviceChange.id,
                        ProcessBy = sessionUser,
                        ProcessedDate = DateTime.UtcNow,
                        RequestText = deviceChange.DeviceChangeRequest?.ChangeRequest,
                        ResponseStatus = BulkChangeStatus.ERROR,
                        ErrorText = statusMessage,
                        ResponseText = $"{CommonStrings.PreflightCheckError} {statusMessage}"
                    });

                    continue;
                }

                deviceChange.ProcessedDate = DateTime.UtcNow;
                deviceChange.ProcessedBy = sessionUser;

                var mobilityDeviceTenant = altaWrxDb.MobilityDevice_Tenant.FirstOrDefault(x => x.MobilityDeviceId == mobilityDevice.id && x.TenantId == permissionManager.Tenant.id);
                if (mobilityDeviceTenant == null)
                {
                    mobilityDeviceTenant = new MobilityDevice_Tenant
                    {
                        MobilityDeviceId = mobilityDevice.id,
                        TenantId = permissionManager.Tenant.id,
                        SiteId = site.id,
                        AccountNumber = revCustomer.RevCustomerId,
                        AccountNumberIntegrationAuthenticationId = revCustomer.IntegrationAuthenticationId,
                        IsActive = true,
                        IsDeleted = false,
                        CreatedDate = DateTime.UtcNow,
                        CreatedBy = sessionUser
                    };

                    if (changeRequestDateIsInvalid && !string.IsNullOrWhiteSpace(changeRequest.CustomerRatePlan))
                    {
                        mobilityDeviceTenant.CustomerRatePlanId = int.Parse(changeRequest.CustomerRatePlan);
                    }

                    if (changeRequestDateIsInvalid && !string.IsNullOrWhiteSpace(changeRequest.CustomerRatePool))
                    {
                        mobilityDeviceTenant.CustomerRatePoolId = int.Parse(changeRequest.CustomerRatePool);
                    }

                    altaWrxDb.MobilityDevice_Tenant.Add(mobilityDeviceTenant);
                }
                else
                {
                    mobilityDeviceTenant.SiteId = site.id;
                    mobilityDeviceTenant.AccountNumber = revCustomer.RevCustomerId;
                    mobilityDeviceTenant.AccountNumberIntegrationAuthenticationId = revCustomer.IntegrationAuthenticationId;

                    if (changeRequestDateIsInvalid && !string.IsNullOrWhiteSpace(changeRequest.CustomerRatePlan))
                    {
                        mobilityDeviceTenant.CustomerRatePlanId = int.Parse(changeRequest.CustomerRatePlan);
                    }

                    if (changeRequestDateIsInvalid && !string.IsNullOrWhiteSpace(changeRequest.CustomerRatePool))
                    {
                        mobilityDeviceTenant.CustomerRatePoolId = int.Parse(changeRequest.CustomerRatePool);
                    }
                    else
                    {
                        mobilityDeviceTenant.CustomerRatePoolId = null;
                    }

                    mobilityDeviceTenant.ModifiedDate = DateTime.UtcNow;
                    mobilityDeviceTenant.ModifiedBy = sessionUser;
                    altaWrxDb.Entry(mobilityDeviceTenant).State = EntityState.Modified;
                }

                // change detail log
                deviceBulkChangeLogRepository.AddMobilityLogEntry(new Amop.Core.Models.DeviceBulkChange.CreateMobilityDeviceBulkChangeLog()
                {
                    BulkChangeId = bulkChangeId,
                    HasErrors = false,
                    LogEntryDescription = "AMOP Customer Assignment",
                    MobilityDeviceChangeId = deviceChange.id,
                    ProcessBy = sessionUser,
                    ProcessedDate = DateTime.UtcNow,
                    RequestText = deviceChange.DeviceChangeRequest?.ChangeRequest,
                    ResponseStatus = BulkChangeStatus.PROCESSED,
                    ResponseText = "ok"
                });

                if (changeRequest.AddCustomerRatePlan && !string.IsNullOrWhiteSpace(changeRequest.CustomerRatePlan))
                {
                    deviceBulkChangeLogRepository.AddMobilityLogEntry(new Amop.Core.Models.DeviceBulkChange.CreateMobilityDeviceBulkChangeLog()
                    {
                        BulkChangeId = bulkChangeId,
                        HasErrors = false,
                        LogEntryDescription = "AMOP Customer Rate Plan Assignment",
                        MobilityDeviceChangeId = deviceChange.id,
                        ProcessBy = sessionUser,
                        ProcessedDate = DateTime.UtcNow,
                        RequestText = deviceChange.DeviceChangeRequest?.ChangeRequest,
                        ResponseStatus = BulkChangeStatus.PROCESSED,
                        ResponseText = "ok"
                    });
                }

                if (changeRequest.AddCustomerRatePlan && !string.IsNullOrWhiteSpace(changeRequest.CustomerRatePool))
                {
                    deviceBulkChangeLogRepository.AddMobilityLogEntry(new Amop.Core.Models.DeviceBulkChange.CreateMobilityDeviceBulkChangeLog()
                    {
                        BulkChangeId = bulkChangeId,
                        HasErrors = false,
                        LogEntryDescription = "AMOP Customer Rate Pool Assignment",
                        MobilityDeviceChangeId = deviceChange.id,
                        ProcessBy = sessionUser,
                        ProcessedDate = DateTime.UtcNow,
                        RequestText = deviceChange.DeviceChangeRequest?.ChangeRequest,
                        ResponseStatus = BulkChangeStatus.PROCESSED,
                        ResponseText = "ok"
                    });
                }

                deviceIdList.Add(deviceChange.DeviceId.ToString());
            }

            await altaWrxDb.SaveChangesAsync();
            if ((!string.IsNullOrWhiteSpace(model.CustomerRatePlan) || !string.IsNullOrWhiteSpace(model.CustomerRatePool)) && permissionManager.OptimizationSettings.OptIntoCrossProviderCustomerOptimization)
            {
                altaWrxDb.usp_UpdateCrossProviderDeviceHistory(string.Empty, string.Join(",", deviceIdList), (int)PortalTypeEnum.Mobility, permissionManager.Tenant.id, model.ServiceProviderId, model.EffectiveDate);
            }
        }

        private Site GetSiteInfo(RevCustomer revCustomer)
        {
            var siteRepository = new SiteRepository(altaWrxDb);
            return siteRepository.GetByRevCustomerId(revCustomer.id);
        }

        private bool UpdateMobilityDeviceHistory(Models.Repositories.MobilityDevice device, MobilityDevice_Tenant deviceTenant, DateTime? effectiveDate)
        {
            var mobilityDeviceHistoryRepository = new MobilityDeviceHistoryRepository(altaWrxDb);

            /*MobilityDeviceHistory mobilityDeviceHistory = device.ToMobilityDeviceHistory(deviceTenant, user.Username, effectiveDate);

            return mobilityDeviceHistoryRepository.InsertMobilityDeviceHistory(mobilityDeviceHistory);*/
            var billingPeriodRepo = new Repositories.BillingPeriod.BillingPeriodRepository(altaWrxDb);
            var billingPeriodIds = new List<int>();
            billingPeriodIds.Add((int)device.BillingPeriodId);
            if (effectiveDate != null)
            {
                billingPeriodIds = billingPeriodRepo.GetBillingPeriodIdsByServiceProviderAndDate(device.ServiceProviderId, (DateTime)effectiveDate);
            }
            var isUpdateMobilityDeviceHistorySuccess = true;
            foreach (var billingPeriodId in billingPeriodIds)
            {
                MobilityDeviceHistory mobilityDeviceHistory = mobilityDeviceHistoryRepository.GetByDeviceTenantAndBillingperiod(deviceTenant.id, billingPeriodId);
                if (mobilityDeviceHistory != null && mobilityDeviceHistory.Id > 0)
                {
                    var updatedDeviceHistory = device.ToUpdatedMobilityDeviceHistory(deviceTenant, user.Username, effectiveDate, mobilityDeviceHistory);
                    if (mobilityDeviceHistoryRepository.UpdateMobilityDeviceHistory(mobilityDeviceHistory, updatedDeviceHistory))
                    {
                        Log.Info(string.Format(LogCommonStrings.UPDATE_MOBILITY_DEVICE_HISTORY_SUCCESS, updatedDeviceHistory.ToString()));
                    }
                    else
                    {
                        Log.Error(string.Format(LogCommonStrings.ERROR_WHILE_UPDATE_MOBILITY_DEVICE_HISTORY, updatedDeviceHistory.ToString()));
                        isUpdateMobilityDeviceHistorySuccess = false;
                    }
                }
            }
            return isUpdateMobilityDeviceHistorySuccess;
        }

        private string MobilityConfigurationQueueFromCustomObjects(IList<CustomObject> customObjectDbList)
        {
            return ValueFromCustomObjects(customObjectDbList, MOBILITY_LINE_CONFIGURATION_QUEUE_NAME);
        }

        private string GetMobilityConfigurationQueueName()
        {
            return MobilityConfigurationQueueFromCustomObjects(GetCustomObjects(permissionManager));
        }

        private List<SelectListItem> AdvancedFilterServiceProviders()
        {
            var serviceProviders = ListHelper.ServiceProviderList(altaWrxDb, PORTAL_TYPE)
                .Select(provider => new SelectListItem { Text = provider.DisplayName, Value = provider.id.ToString() })
                .ToList();
            serviceProviders.Insert(0, new SelectListItem { Text = string.Empty, Value = string.Empty });
            return serviceProviders;
        }

        private List<SelectListItem> AdvancedFilterStatuses()
        {
            var statuses = ListHelper.DeviceStatuses(altaWrxDb, PORTAL_TYPE)
                .Select(status => new SelectListItem { Text = $"{status.DisplayName} ({status.Integration.Name})", Value = status.Status })
                .ToList();
            statuses.Insert(0, new SelectListItem { Text = string.Empty, Value = string.Empty });
            return statuses;
        }

        private DeviceBulkChange BuildBulkChange(int serviceProviderId, List<Mobility_DeviceChange> changes, DeviceChangeType changeType)
        {
            var processedBy = SessionHelper.GetAuditByName(Session);
            var processedDate = DateTime.UtcNow;
            return new DeviceBulkChange
            {
                ChangeRequestTypeId = (int)changeType,
                ServiceProviderId = serviceProviderId,
                TenantId = permissionManager.PermissionFilter.LoggedInTenantId,
                SiteId = GetSiteIdForBulkChange(changes),
                Status = changes.Any(change => !change.IsProcessed) ? BulkChangeStatus.NEW : BulkChangeStatus.PROCESSED,
                CreatedDate = processedDate,
                CreatedBy = processedBy,
                IsActive = true,
                IsDeleted = false,
                Mobility_DeviceChange = changes,
                ProcessedBy = processedBy,
                ProcessedDate = processedDate,
                ModifiedBy = processedBy,
                ModifiedDate = processedDate
            };
        }

        [HttpPost]
        public ActionResult ValidateStaticIPSettings(int serviceProviderId)
        {
            if (!permissionManager.UserCanCreate(Session, ModuleEnum.Mobility))
            {
                return new HttpStatusCodeResult(HttpStatusCode.Forbidden);
            }

            var serviceProviderSettingRepository = new ServiceProviderSettingRepository(altaWrxDb);
            var pdpNames = serviceProviderSettingRepository.GetByServiceProviderId("PDPName", serviceProviderId).FirstOrDefault();
            var APNs = serviceProviderSettingRepository.GetByServiceProviderId("StaticIPAPN", serviceProviderId).FirstOrDefault();

            return Json(new
            {
                isValid = !string.IsNullOrEmpty(pdpNames.SettingValue) && !string.IsNullOrEmpty(APNs.SettingValue)
            });
        }

        public async Task<bool> IsTelegenceBillingAccountNumberValid(int serviceProviderId, string billingAccountNumber)
        {
            var isValid = true;
            var mobilityDeviceRepository = new MobilityDeviceRepository(altaWrxDb);
            var banFromMobilityDevice = mobilityDeviceRepository.GetBillingAccountNumberFromMobilityDevice(serviceProviderId, billingAccountNumber);
            if (string.IsNullOrWhiteSpace(banFromMobilityDevice) && IsProduction)
            {
                var telegenceAuthentication = altaWrxDb.usp_Telegence_Get_AuthenticationByProviderId(serviceProviderId).Select(auth => auth.ToTelegenceAuthentication()).FirstOrDefault();
                if (telegenceAuthentication != null)
                {
                    var telegenceAPIClient = new TelegenceAPIClient(new SingletonHttpClientFactory(), new HttpRequestFactory(), telegenceAuthentication, IsProduction, string.Empty, null);
                    var banFromAPI = await telegenceAPIClient.GetBillingAccountNumber(billingAccountNumber, URLConstants.TELEGENCE_BAN_DETAIL_GET_URL);
                    if (string.IsNullOrWhiteSpace(banFromAPI?.BusinessAccount?.Fan))
                    {
                        isValid = false;
                    }
                }
                else
                {
                    isValid = false;
                }
            }

            return isValid;
        }

        public FileContentResult DownloadServiceActivationTemplate(int qualificationId = 0, bool isAgent = false)
        {
            var fileName = isAgent ? "ServiceActivationTemplateAgent.xlsx" : "BulkActivationUploadTemplateV2.xlsx";

            var file = FileSystem.ReadAllBytes(Server.MapPath(isAgent ? "~/Content/BulkActivationUploadTemplateAgent.xlsx" : "~/Content/BulkActivationUploadTemplateV2.xlsx"));

            var qualificationRepo = new QualificationRepository(altaWrxDb);
            var qual = qualificationRepo.GetByIdWithAddresses(qualificationId);

            if (qual == null)
            {
                return File(file, ExcelContentType, $"{fileName}");
            }

            return File(ExcelUtilities.ExportQualificationTemplate(file, qual), ExcelContentType, $"{fileName}");
        }
    }
}
