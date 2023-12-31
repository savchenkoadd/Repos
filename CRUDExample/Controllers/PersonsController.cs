﻿using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Rotativa.AspNetCore;
using Rotativa.AspNetCore.Options;
using ServiceContracts;
using ServiceContracts.DTO;
using ServiceContracts.Enums;

namespace CRUDExample.Controllers
{
	[Route("[controller]")]
	public class PersonsController : Controller
	{
		private readonly IPersonsService _personsService;
		private readonly ICountriesService _countriesService;

		public PersonsController(IPersonsService personsService, ICountriesService countriesService)
		{
			_personsService = personsService;
			_countriesService = countriesService;
		}

		//Url: persons/index
		[Route("[action]")]
		[Route("/")]
		public async Task<IActionResult> Index(string searchBy, string? searchString, string sortBy = nameof(PersonResponse.PersonName), SortOrderOptions sortOrder = SortOrderOptions.ASC)
		{
			//Search
			ViewBag.SearchFields = new Dictionary<string, string>()
	  {
		{ nameof(PersonResponse.PersonName), "Person Name" },
		{ nameof(PersonResponse.Email), "Email" },
		{ nameof(PersonResponse.DateOfBirth), "Date of Birth" },
		{ nameof(PersonResponse.Gender), "Gender" },
		{ nameof(PersonResponse.CountryID), "Country" },
		{ nameof(PersonResponse.Address), "Address" }
	  };
			List<PersonResponse> persons = await _personsService.GetFilteredPersons(searchBy, searchString);
			ViewBag.CurrentSearchBy = searchBy;
			ViewBag.CurrentSearchString = searchString;

			//Sort
			List<PersonResponse> sortedPersons = await _personsService.GetSortedPersons(persons, sortBy, sortOrder);
			ViewBag.CurrentSortBy = sortBy;
			ViewBag.CurrentSortOrder = sortOrder.ToString();

			return View(sortedPersons);
		}

		[Route("[action]")]
		[HttpGet]
		public async Task<IActionResult> Create()
		{
			List<CountryResponse> countries = await _countriesService.GetAllCountries();
			ViewBag.Countries = countries.Select(temp =>
			  new SelectListItem() { Text = temp.CountryName, Value = temp.CountryID.ToString() }
			);

			return View();
		}

		[HttpPost]
		[Route("[action]")]
		public async Task<IActionResult> Create(PersonAddRequest personAddRequest)
		{
			if (!ModelState.IsValid)
			{
				List<CountryResponse> countries = await _countriesService.GetAllCountries();
				ViewBag.Countries = countries.Select(temp =>
				new SelectListItem() { Text = temp.CountryName, Value = temp.CountryID.ToString() });

				ViewBag.Errors = ModelState.Values.SelectMany(v => v.Errors).Select(e => e.ErrorMessage).ToList();
				return View();
			}

			PersonResponse personResponse = await _personsService.AddPerson(personAddRequest);

			return RedirectToAction("Index", "Persons");
		}

		[HttpGet]
		[Route("[action]")]
		public async Task<IActionResult> PersonsPDF(PersonAddRequest personAddRequest)
		{
			var persons = await _personsService.GetAllPersons();

			return new ViewAsPdf(viewName: "PersonsPDF", persons, ViewData)
			{
				PageMargins = new Margins() { Top = 20, Bottom = 20, Left = 30, Right = 30 },
				PageOrientation = Orientation.Landscape
			};
		}

		[HttpGet]
		[Route("[action]")]
		public async Task<IActionResult> PersonsCSV()
		{
			var memoryStream = await _personsService.GetPersonsCSV();

			return File(memoryStream, contentType: "application/octet-stream", fileDownloadName: "persons.csv");
		}

		[HttpGet]
		[Route("[action]")]
		public async Task<IActionResult> PersonsExcel()
		{
			var memoryStream = await _personsService.GetPersonsExcel();

			return File(memoryStream, contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileDownloadName: "persons.xlsx");
		}

		[HttpGet]
		[Route("[action]/{personID}")]
		public async Task<IActionResult> Edit(Guid personID)
		{
			PersonResponse? personResponse = await _personsService.GetPersonByPersonID(personID);
			if (personResponse == null)
			{
				return RedirectToAction("Index");
			}

			PersonUpdateRequest personUpdateRequest = personResponse.ToPersonUpdateRequest();

			List<CountryResponse> countries = await _countriesService.GetAllCountries();
			ViewBag.Countries = countries.Select(temp =>
			new SelectListItem() { Text = temp.CountryName, Value = temp.CountryID.ToString() });

			return View(personUpdateRequest);
		}


		[HttpPost]
		[Route("[action]/{personID}")]
		public async Task<IActionResult> Edit(PersonUpdateRequest personUpdateRequest)
		{
			PersonResponse? personResponse = await _personsService.GetPersonByPersonID(personUpdateRequest.PersonID);

			if (personResponse == null)
			{
				return RedirectToAction("Index");
			}

			if (ModelState.IsValid)
			{
				PersonResponse updatedPerson = await _personsService.UpdatePerson(personUpdateRequest);
				return RedirectToAction("Index");
			}
			else
			{
				List<CountryResponse> countries = await _countriesService.GetAllCountries();
				ViewBag.Countries = countries.Select(temp =>
				new SelectListItem() { Text = temp.CountryName, Value = temp.CountryID.ToString() });

				ViewBag.Errors = ModelState.Values.SelectMany(v => v.Errors).Select(e => e.ErrorMessage).ToList();
				return View(personResponse.ToPersonUpdateRequest());
			}
		}


		[HttpGet]
		[Route("[action]/{personID}")]
		public async Task<IActionResult> Delete(Guid? personID)
		{
			PersonResponse? personResponse = await _personsService.GetPersonByPersonID(personID);
			if (personResponse == null)
				return RedirectToAction("Index");

			return View(personResponse);
		}

		[HttpPost]
		[Route("[action]/{personID}")]
		public async Task<IActionResult> Delete(PersonUpdateRequest personUpdateResult)
		{
			PersonResponse? personResponse = await _personsService.GetPersonByPersonID(personUpdateResult.PersonID);
			if (personResponse == null)
				return RedirectToAction("Index");

			await _personsService.DeletePerson(personUpdateResult.PersonID);
			return RedirectToAction("Index");
		}
	}
}
