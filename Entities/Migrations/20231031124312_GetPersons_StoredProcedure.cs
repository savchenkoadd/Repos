using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Entities.Migrations
{
	public partial class GetPersons_StoredProcedure : Migration
	{
		protected override void Up(MigrationBuilder migrationBuilder)
		{
			string asp_GetAllPersons = @"
                CREATE PROCEDURE [dbo].[GetAllPersons]
                AS BEGIN
					SELECT * from [dbo].[Persons]
				END
            ";

			migrationBuilder.Sql(asp_GetAllPersons);
		}

		protected override void Down(MigrationBuilder migrationBuilder)
		{
			string asp_GetAllPersons = @"
                DROP PROCEDURE [dbo].[GetAllPersons]
            ";

			migrationBuilder.Sql(asp_GetAllPersons);
		}
	}
}
