using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

public class StorageDetail
{
    [Key]
    public string BarCodeValue { get; set; }

    public string PatientId { get; set; }
    public string SiteNo { get; set; }
    public string VisitName { get; set; }
    public DateTime? DOBSCollection { get; set; }
    public TimeSpan? TOBSCollection { get; set; }
    public string TypeOfSample { get; set; }
    public string Cryobox { get; set; }
    public int? CryoboxWellPosition { get; set; }
    public string RequisitionId { get; set; }
    public string Remarks { get; set; }
    public string ReceivedCondition { get; set; }
    public string ProjectCode { get; set; }
    public string Status { get; set; }
    public string Location { get; set; }
}

public class TRF_RegistrationInfo
{
    [Key]
    public string RequisitionId { get; set; }

    public long SystemId { get; set; }
    public string PatientId { get; set; }
    public string SiteNo { get; set; }
    public string VisitName { get; set; }
    public string Gender { get; set; }
    public string PatientInitials { get; set; }
    public string PatientName { get; set; }
    public string SubjectId { get; set; }
    public DateTime? CollectionDate { get; set; }
    public DateTime? ProcessingDate { get; set; }
    public bool? RegistrationStatus { get; set; }
    public bool? RegistrationCheckStatus { get; set; }
    public bool? ResultCheckStatus { get; set; }
    public bool? ApprovedStatus { get; set; }
    public bool? HoldStatus { get; set; }
    public bool? IsLock { get; set; }
    public DateTime? CreatedOn { get; set; }
    public bool? ReportGeneratedStatus { get; set; }
    public string RegField3 { get; set; }
    public decimal? Height { get; set; }
    public decimal? Weight { get; set; }
    public long? AgeValue { get; set; }
    public DateTime? DateOfBirth { get; set; }
    public string StudyId { get; set; }
    public string ProjectCode { get; set; }
    public DateTime? ReceivedDateTime { get; set; }
}

public class TRF_Reg_BarCode
{
    [Key]
    public string BarCodeValue { get; set; }

    public string ProjectCode { get; set; }
    public long SystemId { get; set; }
    public string RequisitionId { get; set; }
    public string Remarks { get; set; }
    public string ReceivedCondition { get; set; }
    public string Status { get; set; }
    public DateTime? CreatedDateTime { get; set; }
    public DateTime? CollectionDate { get; set; }
    public string CustomField1 { get; set; }
}

public class AuditHistory
{
    [Key]
    [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
    public long SystemId { get; set; }  // Use 'long' for bigint

    public string LoginId { get; set; }
    public string ProjectCode { get; set; }
    public string PageName { get; set; }
    public string ModuleName { get; set; }
    public string UserRemarks { get; set; }
    public DateTime? CreatedDateTime { get; set; }
    public string SystemRemarks { get; set; }
    public long TransId { get; set; }
}