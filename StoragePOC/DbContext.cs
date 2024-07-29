using System.Data.Entity;

public class LIMSDevContext : DbContext
{
    public DbSet<TRF_RegistrationInfo> TRF_RegistrationInfos { get; set; }
    public DbSet<TRF_Reg_BarCode> TRF_Reg_BarCodes { get; set; }
    public DbSet<StorageDetail> StorageDetails { get; set; }
    public DbSet<AuditHistory> AuditHistorys { get; set; }

    public LIMSDevContext() : base("name=LIMSDevContext")
    {
        //if it is set then EF will not create DB if does not exist 
        Database.SetInitializer<LIMSDevContext>(null);
    }

    protected override void OnModelCreating(DbModelBuilder modelBuilder)
    {
        base.OnModelCreating(modelBuilder);
        modelBuilder.Entity<TRF_RegistrationInfo>().ToTable("TRF_RegistrationInfos");
        modelBuilder.Entity<TRF_Reg_BarCode>().ToTable("TRF_Reg_Barcodes");
        modelBuilder.Entity<StorageDetail>().ToTable("StorageDetail");
        modelBuilder.Entity<AuditHistory>().ToTable("AuditHistory");
    }
}
