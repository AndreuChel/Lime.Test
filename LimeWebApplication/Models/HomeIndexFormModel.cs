using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace LimeTest.Models
{
    /// <summary>
    /// Модель для формы на главной странице
    /// </summary>
    public class HomeIndexFormModel
    {
        [Display(Name = "Эл.почта")]
        [Required]
        [EmailAddress(ErrorMessage = "Неверный адрес эл.почты")]
        public string Email
        { get; set; }

        [Display(Name = "Начало периода")]
        public DateTime? StartDate { get; set; }

        [Display(Name = "Конец периода")]
        public DateTime? EndDate { get; set; }
    }
}