using LimeTestApp.Data.NorthwindDb.Entities;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LimeTestApp.Data.NorthwindDb
{
    public interface INorthwindContext : IDisposable
    {
        DbSet<Category> Category { get; set; }
        DbSet<Customer> Customer { get; set; }
        DbSet<Employee> Employee { get; set; }
        DbSet<Order> Order { get; set; }
        DbSet<OrderDetail> OrderDetail { get; set; }
        DbSet<Product> Product { get; set; }
        DbSet<Shipper> Shipper { get; set; }
        DbSet<Supplier> Supplier { get; set; }
    }
}
