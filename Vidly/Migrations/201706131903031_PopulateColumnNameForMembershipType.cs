namespace Vidly.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class PopulateColumnNameForMembershipType : DbMigration
    {
        public override void Up()
        {
            Sql("UPDATE MembershipTypes SET Name = 'Bronze' WHERE Id = 1");
            Sql("UPDATE MembershipTypes SET Name = 'Silver' WHERE Id = 2");
            Sql("UPDATE MembershipTypes SET Name = 'Gold' WHERE Id = 3");
            Sql("UPDATE MembershipTypes SET Name = 'Platinum' WHERE Id = 4");
        }
        
        public override void Down()
        {
        }
    }
}
