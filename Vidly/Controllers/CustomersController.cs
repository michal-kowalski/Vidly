using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
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
			var customers = _context.Customers.ToList();

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
			var customer = _context.Customers.SingleOrDefault(c => c.Id == id);
			if (customer == null)
				return HttpNotFound();

			var customers = new Customer();

			var ViewModel = new RandomMovieViewModel
			{
				Customer = customer
			};

			
			return View(ViewModel);
		}

	}
}