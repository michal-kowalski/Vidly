﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace Vidly.Models
{
	public class Movie
	{
		public int Id { get; set; }
        [Required]
        [MaxLength(512)]
		public string Name { get; set; }
		public string Description { get; set; }
		public double Rating { get; set; }
	}
}