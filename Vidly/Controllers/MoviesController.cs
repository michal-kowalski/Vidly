using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Vidly.Models;
using Vidly.ViewModels;

namespace Vidly.Controllers
{
	public class MoviesController : Controller
	{
        private ApplicationDbContext _context;
        public MoviesController()
        {
            _context = new ApplicationDbContext();
        }

        // GET: Movies
        public ActionResult Index()
		{
            var movies = _context.Movies.ToList();
			/*var movies = new List<Movie>
			{
				new Movie { Name = "Matrix", Id = 0, Description = "weqeqeq", Rating = 5.0 },
				new Movie { Name = "Matrix 2", Id = 1, Description = "weqeqeq", Rating = 4.5 },
				new Movie { Name = "Matrix 3", Id = 2, Description = "weqeqeq", Rating = 3.6 }

			};*/


			var viewModel = new RandomMovieViewModel
			{
				Movie = movies
			};

			return View(viewModel);
		}


		// GET: Movies/Random
		public ActionResult Random()
		{
			var movie = new Movie() { Name = "Shrek!", Id = 1 };
			var customers = new List<Customer>
			{
				new Customer {Name = "Customer 1", Id = 1 },
				new Customer {Name = "Customer 2", Id = 3 },
				new Customer {Name = "Customer 3", Id = 3 }
			};

			var viewModel = new RandomMovieViewModel
			{
				Customers = customers,
                OneMovie = movie
                
			};

			return View(viewModel);
		}

		public ActionResult Edit(int id)
		{
			return Content("id=" + id);
		}

		public ActionResult ByReleaseDate(int year, int month)
		{
			return Content(year + "/" + month);
		}

		public ActionResult MoviesDetails(int id)
		{
            var movie = _context.Movies.SingleOrDefault(c => c.Id == id);
            var viewModel = new RandomMovieViewModel
            {
                OneMovie = movie
            };

			return View(viewModel);
		}
	}


}