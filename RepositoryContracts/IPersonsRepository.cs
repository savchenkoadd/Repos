using Entities;
using System.Linq.Expressions;

namespace RepositoryContracts
{
	public interface IPersonsRepository
	{
		Task<Person> AddPerson(Person person);

		Task<IEnumerable<Person>> GetAllPersons();

		Task<IEnumerable<Person>> GetFilteredPersons(Expression<Func<Person, bool>> predicate);

		Task<Person?> GetPersonById(Guid id);

		Task<Person?> GetPersonByPersonName(string personName);

		Task<bool> DeletePersonByPersonId(Guid id);

		Task<Person> UpdatePerson(Person person);
	}
}
