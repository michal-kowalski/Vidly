namespace Vidly.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class AddedGenresTableSecondAttemp1 : DbMigration
    {
        public override void Up()
        {
            Sql("INSERT INTO genres (Id, Name) VALUES (1, 'Comedy')");
            Sql("INSERT INTO genres (Id, Name) VALUES (2, 'Drama')");
            Sql("INSERT INTO genres (Id, Name) VALUES (3, 'Mystery')");
            Sql("INSERT INTO genres (Id, Name) VALUES (4, 'Anime')");
            Sql("INSERT INTO genres (Id, Name) VALUES (5, 'Fantasy')");

        }
        
        public override void Down()
        {
        }
    }
}
