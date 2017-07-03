using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data.Entity;
using Vidly.Models;
using Vidly.ViewModels;

namespace Vidly.Controllers
{
	public class CustomersController : Controller
	{
		private ApplicationDbContext _context;
		public CustomersController()
		{
			_context = new ApplicationDbContext();
		}
		protected override void Dispose(bool disposing)
		{
			_context.Dispose();
		}
										   // GET: Customers
		public ActionResult Index()
		{
			var movie = new Movie() { Name = "Shrek!", Id = 1 };
			var customers = _context.Customers.Include(c=>c.MembershipType).ToList();

			/*var customers = new List<Customer>
			{
				new Customer {Name = "John Smith", Id = 0 },
				new Customer {Name = "Jon Doe", Id = 1 },
				new Customer {Name = "Jane Doe", Id = 2 },
				new Customer {Name = "Customer 4", Id = 3 }
			};*/

			var viewModel = new RandomMovieViewModel
			{

				Customers = customers
			};

			return View(viewModel);
		}

		public ActionResult CustomersDetails(int id)
		{
			var customer = _context.Customers.Include(c=>c.MembershipType).SingleOrDefault(c => c.Id == id);
			if (customer == null)
				return HttpNotFound();

			var ViewModel = new RandomMovieViewModel
			{
				Customer = customer
			};

			
			return View(ViewModel);
		}

		public ActionResult New()
		{
			var membershipTypes = _context.MembershipTypes.ToList();
			var viewModel = new CustomerFormViewModel
			{
				MembershipTypes = membershipTypes
			};
			return View("CustomerForm", viewModel);
		}

		[HttpPost]
		public ActionResult Create(Customer customer)
		{
			_context.Customers.Add(customer);
			_context.SaveChanges();
			return RedirectToAction("Index", "Customers");
		}

		public ActionResult Edit(int id)
		{
			var customer = _context.Customers.SingleOrDefault(c => c.Id == id);
			if (customer == null)
				return HttpNotFound();

			var viewModel = new CustomerFormViewModel
			{
				Customer = customer,
				MembershipTypes = _context.MembershipTypes.ToList()
			};
			return View("CustomerForm", viewModel);

		}

	}
}