using Entities;
using ServiceContracts.DTO;
using ServiceContracts;
using Services.Helpers;
using ServiceContracts.Enums;
using Microsoft.EntityFrameworkCore;
using CsvHelper;
using System.Globalization;
using CsvHelper.Configuration;
using OfficeOpenXml;
using RepositoryContracts;

namespace Services
{
	public class PersonsService : IPersonsService
	{
		//private field
		private readonly IPersonsRepository _repository;
		private readonly ICountriesService _countriesService;

		//constructor
		public PersonsService(
				IPersonsRepository repository,
				ICountriesService countriesService
			)
		{
			_repository = repository;
			_countriesService = countriesService;
		}


		private async Task<PersonResponse> ConvertPersonToPersonResponse(Person person)
		{
			PersonResponse personResponse = person.ToPersonResponse();
			personResponse.Country = (await _countriesService.GetCountryByCountryID(person.CountryID))?.CountryName;
			return personResponse;
		}

		public async Task<PersonResponse> AddPerson(PersonAddRequest? personAddRequest)
		{
			//check if PersonAddRequest is not null
			if (personAddRequest == null)
			{
				throw new ArgumentNullException(nameof(personAddRequest));
			}

			//Model validation
			ValidationHelper.ModelValidation(personAddRequest);

			//convert personAddRequest into Person type
			Person person = personAddRequest.ToPerson();

			//generate PersonID
			person.PersonID = Guid.NewGuid();

			//add person object to persons list
			await _repository.Persons.AddAsync(person);
			await _repository.SaveChangesAsync();

			//convert the Person object into PersonResponse type
			return await ConvertPersonToPersonResponse(person);
		}


		public async Task<List<PersonResponse>> GetAllPersons()
		{
			var people = await _repository.Persons.Include("Country").ToListAsync();

			return people.Select(temp => temp.ToPersonResponse()).ToList();
		}

		public async Task<PersonResponse?> GetPersonByPersonID(Guid? personID)
		{
			if (personID == null)
				return null;

			Person? person = await _repository.Persons.FirstOrDefaultAsync(temp => temp.PersonID == personID);
			if (person == null)
				return null;

			return person.ToPersonResponse();
		}

		public async Task<List<PersonResponse>> GetFilteredPersons(string searchBy, string? searchString)
		{
			List<PersonResponse> allPersons = await GetAllPersons();
			List<PersonResponse> matchingPersons = allPersons;

			if (string.IsNullOrEmpty(searchBy) || string.IsNullOrEmpty(searchString))
				return matchingPersons;

			switch (searchBy)
			{
				case nameof(PersonResponse.PersonName):
					matchingPersons = allPersons.Where(temp =>
					(!string.IsNullOrEmpty(temp.PersonName) ?
					temp.PersonName.Contains(searchString, StringComparison.OrdinalIgnoreCase) : true)).ToList();
					break;

				case nameof(PersonResponse.Email):
					matchingPersons = allPersons.Where(temp =>
					(!string.IsNullOrEmpty(temp.Email) ?
					temp.Email.Contains(searchString, StringComparison.OrdinalIgnoreCase) : true)).ToList();
					break;


				case nameof(PersonResponse.DateOfBirth):
					matchingPersons = allPersons.Where(temp =>
					(temp.DateOfBirth != null) ?
					temp.DateOfBirth.Value.ToString("dd MMMM yyyy").Contains(searchString, StringComparison.OrdinalIgnoreCase) : true).ToList();
					break;

				case nameof(PersonResponse.Gender):
					matchingPersons = allPersons.Where(temp =>
					(!string.IsNullOrEmpty(temp.Gender) ?
					temp.Gender.Contains(searchString, StringComparison.OrdinalIgnoreCase) : true)).ToList();
					break;

				case nameof(PersonResponse.CountryID):
					matchingPersons = allPersons.Where(temp =>
					(!string.IsNullOrEmpty(temp.Country) ?
					temp.Country.Contains(searchString, StringComparison.OrdinalIgnoreCase) : true)).ToList();
					break;

				case nameof(PersonResponse.Address):
					matchingPersons = allPersons.Where(temp =>
					(!string.IsNullOrEmpty(temp.Address) ?
					temp.Address.Contains(searchString, StringComparison.OrdinalIgnoreCase) : true)).ToList();
					break;

				default: matchingPersons = allPersons; break;
			}
			return matchingPersons;
		}

		public async Task<List<PersonResponse>> GetSortedPersons(List<PersonResponse> allPersons, string sortBy, SortOrderOptions sortOrder)
		{
			if (string.IsNullOrEmpty(sortBy))
				return allPersons;

			List<PersonResponse> sortedPersons = (sortBy, sortOrder) switch
			{
				(nameof(PersonResponse.PersonName), SortOrderOptions.ASC) => allPersons.OrderBy(temp => temp.PersonName, StringComparer.OrdinalIgnoreCase).ToList(),

				(nameof(PersonResponse.PersonName), SortOrderOptions.DESC) => allPersons.OrderByDescending(temp => temp.PersonName, StringComparer.OrdinalIgnoreCase).ToList(),

				(nameof(PersonResponse.Email), SortOrderOptions.ASC) => allPersons.OrderBy(temp => temp.Email, StringComparer.OrdinalIgnoreCase).ToList(),

				(nameof(PersonResponse.Email), SortOrderOptions.DESC) => allPersons.OrderByDescending(temp => temp.Email, StringComparer.OrdinalIgnoreCase).ToList(),

				(nameof(PersonResponse.DateOfBirth), SortOrderOptions.ASC) => allPersons.OrderBy(temp => temp.DateOfBirth).ToList(),

				(nameof(PersonResponse.DateOfBirth), SortOrderOptions.DESC) => allPersons.OrderByDescending(temp => temp.DateOfBirth).ToList(),

				(nameof(PersonResponse.Age), SortOrderOptions.ASC) => allPersons.OrderBy(temp => temp.Age).ToList(),

				(nameof(PersonResponse.Age), SortOrderOptions.DESC) => allPersons.OrderByDescending(temp => temp.Age).ToList(),

				(nameof(PersonResponse.Gender), SortOrderOptions.ASC) => allPersons.OrderBy(temp => temp.Gender, StringComparer.OrdinalIgnoreCase).ToList(),

				(nameof(PersonResponse.Gender), SortOrderOptions.DESC) => allPersons.OrderByDescending(temp => temp.Gender, StringComparer.OrdinalIgnoreCase).ToList(),

				(nameof(PersonResponse.Country), SortOrderOptions.ASC) => allPersons.OrderBy(temp => temp.Country, StringComparer.OrdinalIgnoreCase).ToList(),

				(nameof(PersonResponse.Country), SortOrderOptions.DESC) => allPersons.OrderByDescending(temp => temp.Country, StringComparer.OrdinalIgnoreCase).ToList(),

				(nameof(PersonResponse.Address), SortOrderOptions.ASC) => allPersons.OrderBy(temp => temp.Address, StringComparer.OrdinalIgnoreCase).ToList(),

				(nameof(PersonResponse.Address), SortOrderOptions.DESC) => allPersons.OrderByDescending(temp => temp.Address, StringComparer.OrdinalIgnoreCase).ToList(),

				(nameof(PersonResponse.ReceiveNewsLetters), SortOrderOptions.ASC) => allPersons.OrderBy(temp => temp.ReceiveNewsLetters).ToList(),

				(nameof(PersonResponse.ReceiveNewsLetters), SortOrderOptions.DESC) => allPersons.OrderByDescending(temp => temp.ReceiveNewsLetters).ToList(),

				_ => allPersons
			};

			return sortedPersons;
		}

		public async Task<PersonResponse> UpdatePerson(PersonUpdateRequest? personUpdateRequest)
		{
			if (personUpdateRequest == null)
				throw new ArgumentNullException(nameof(personUpdateRequest));

			//validation
			ValidationHelper.ModelValidation(personUpdateRequest);

			//get matching person object to update
			Person? matchingPerson = await _repository.Persons.FirstOrDefaultAsync(temp => temp.PersonID == personUpdateRequest.PersonID);
			if (matchingPerson == null)
			{
				throw new ArgumentException("Given person id doesn't exist");
			}

			//update all details
			matchingPerson.PersonName = personUpdateRequest.PersonName;
			matchingPerson.Email = personUpdateRequest.Email;
			matchingPerson.DateOfBirth = personUpdateRequest.DateOfBirth;
			matchingPerson.Gender = personUpdateRequest.Gender.ToString();
			matchingPerson.CountryID = personUpdateRequest.CountryID;
			matchingPerson.Address = personUpdateRequest.Address;
			matchingPerson.ReceiveNewsLetters = personUpdateRequest.ReceiveNewsLetters;

			await _repository.SaveChangesAsync();

			return await ConvertPersonToPersonResponse(matchingPerson);
		}

		public async Task<bool> DeletePerson(Guid? personID)
		{
			if (personID == null)
			{
				throw new ArgumentNullException(nameof(personID));
			}

			Person? person = await _repository.Persons.FirstOrDefaultAsync(temp => temp.PersonID == personID);
			if (person == null)
				return false;

			_repository.Persons.Remove(await _repository.Persons.FirstAsync(temp => temp.PersonID == personID));
			await _repository.SaveChangesAsync();

			return true;
		}

		public async Task<MemoryStream> GetPersonsCSV()
		{
			MemoryStream result = new MemoryStream();
			StreamWriter writer = new StreamWriter(result);

			CsvConfiguration csvConfiguration = new CsvConfiguration(cultureInfo: CultureInfo.InvariantCulture);
			CsvWriter csvWriter = new CsvWriter(writer, csvConfiguration, leaveOpen: true);

			csvWriter.WriteField(nameof(PersonResponse.PersonName));
			csvWriter.WriteField(nameof(PersonResponse.Email));

			csvWriter.NextRecord();

			foreach (PersonResponse item in await GetAllPersons())
			{
				csvWriter.WriteField(item.PersonName);
				csvWriter.WriteField(item.Email);
				csvWriter.NextRecord();
				await csvWriter.FlushAsync();
			}

			result.Position = 0;

			return result;
		}
		  
		public async Task<MemoryStream> GetPersonsExcel()
		{
			MemoryStream memoryStream = new MemoryStream();

			using (ExcelPackage excelPackage = new ExcelPackage(memoryStream))
			{
				ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("PersonsSheet1"); 

				worksheet.Cells["A1"].Value = "Person Name";
				worksheet.Cells["B1"].Value = "Email";


					int row = 2;

				foreach (var item in await GetAllPersons())
				{
					worksheet.Cells[row, 1].Value = item.PersonName;
					worksheet.Cells[row, 2].Value = item.Email;

					row++;
				}

				using (ExcelRange headerCells = worksheet.Cells[$"A1:H{row}"])
				{
					headerCells.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
					headerCells.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
					headerCells.Style.Font.Bold = true;
				}

				worksheet.Cells[$"A1:B{row}"].AutoFitColumns();

				await excelPackage.SaveAsync();
			}

			memoryStream.Position = 0;

			return memoryStream;
		}
	}
}
 