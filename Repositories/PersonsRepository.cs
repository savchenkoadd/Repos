using Entities;
using Microsoft.EntityFrameworkCore;
using RepositoryContracts;
using System;
using System.Linq.Expressions;

namespace Repositories
{
	public class PersonsRepository : IPersonsRepository
	{
		private readonly ApplicationDbContext _db;

        public PersonsRepository(ApplicationDbContext db)
        {
            _db = db;
        }

        public async Task<Person> AddPerson(Person person)
		{
			await _db.Persons.AddAsync(person);
			await _db.SaveChangesAsync();

			return person;
		}

		public async Task<bool> DeletePersonByPersonId(Guid id)
		{
			_db.Persons.RemoveRange(_db.Persons.Where(x => x.PersonID == id));
			var rowsAffected = await _db.SaveChangesAsync();
			
			return rowsAffected > 0;
		}

		public async Task<IEnumerable<Person>> GetAllPersons()
		{
			return await _db.Persons.Include("Country").ToListAsync();
		}

		public async Task<IEnumerable<Person>> GetFilteredPersons(Expression<Func<Person, bool>> predicate)
		{
			return await _db.Persons.Include("Country").Where(predicate).ToListAsync();
		}

		public async Task<Person?> GetPersonById(Guid id)
		{
			return await _db.Persons.Include("Country").Where(x => x.PersonID == id).FirstOrDefaultAsync();
		}

		public Task<Person> GetPersonByPersonName(string personName)
		{
			throw new NotImplementedException();
		}

		public async Task<Person> UpdatePerson(Person person)
		{
			Person? matchingPerson = await _db.Persons.FirstOrDefaultAsync(x => x.PersonID == person.PersonID);

			if (matchingPerson == null)
			{
				return person;
			}

			matchingPerson.PersonID = person.PersonID;
			matchingPerson.Gender = person.Gender;
			matchingPerson.Address = person.Address;
			matchingPerson.Country = person.Country;
			matchingPerson.CountryID = person.CountryID;
			matchingPerson.DateOfBirth = person.DateOfBirth;
			matchingPerson.Email = person.Email;
			matchingPerson.TIN = person.TIN;
			matchingPerson.PersonName = person.PersonName;
			matchingPerson.ReceiveNewsLetters = person.ReceiveNewsLetters;

			await _db.SaveChangesAsync();

			return matchingPerson;
		}
	}
}
