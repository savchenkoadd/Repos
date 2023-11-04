using System;
using System.Collections.Generic;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;

namespace Entities
{
	public class ApplicationDbContext : DbContext
	{
		public ApplicationDbContext(DbContextOptions options) : base(options)
		{
		}

		public virtual DbSet<Country> Countries { get; set; }
		public virtual DbSet<Person> Persons { get; set; }

		protected override void OnModelCreating(ModelBuilder modelBuilder)
		{
			base.OnModelCreating(modelBuilder);

			modelBuilder.Entity<Country>().ToTable("Countries");
			modelBuilder.Entity<Person>().ToTable("Persons");

			//Seed to Countries
			string countriesJson = File.ReadAllText("countries.json");
			List<Country> countries = System.Text.Json.JsonSerializer.Deserialize<List<Country>>(countriesJson);

			foreach (Country country in countries)
				modelBuilder.Entity<Country>().HasData(country);


			//Seed to Persons
			string personsJson = File.ReadAllText("persons.json");
			List<Person> persons = System.Text.Json.JsonSerializer.Deserialize<List<Person>>(personsJson);
			  
			foreach (Person person in persons)
				modelBuilder.Entity<Person>().HasData(person);

			//Fluent API
			modelBuilder.Entity<Person>().Property(temp => temp.TIN)
				.HasColumnName("TaxIdentificationNumber")
				.HasColumnType("varchar(8)")
				.HasDefaultValue("0");

			modelBuilder.Entity<Person>()
				.HasCheckConstraint("CHK_TIN", "len([TaxIdentificationNumber]) = 8");
		}

		public List<Person> sp_GetAllPersons()
		{
			return Persons.FromSqlRaw("EXECUTE [dbo].[GetAllPersons]").ToList();
		}

		public int sp_InsertPerson(Person person)
		{
			SqlParameter[] parameters = new SqlParameter[]
			{
				new SqlParameter("@PersonId", person.PersonID),
				new SqlParameter("@PersonName", person.PersonName),
				new SqlParameter("@Email", person.Email),
				new SqlParameter("@Gender", person.Gender),
				new SqlParameter("@Address", person.Address),
				new SqlParameter("@DateOfBirth", person.DateOfBirth),
				new SqlParameter("@ReceiveNewsLetters", person.ReceiveNewsLetters),
				new SqlParameter("@CountryID", person.CountryID),
			};

			return Database.ExecuteSqlRaw(@"EXECUTE [dbo].[InsertPerson] @PersonId, @PersonName, @Email, @DateOfBirth, @Gender, @CountryId, @Address, @ReceiveNewsLetters", parameters);
		}
	}
}
