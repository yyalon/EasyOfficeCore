using EasyOffice.Enums;
using EasyOffice.Interfaces;
using EasyOffice.Providers.NPOI;
using EasyOffice.Services;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Extensions
{
    public static class EasyOfficeExtensions
    {
        public static IServiceCollection UseEasyOffice(this IServiceCollection services, SolutionEnum solutionEnum)
        {
            switch (solutionEnum)
            {
                case SolutionEnum.NPOI:
                    services.AddTransient<IWordExportProvider, WordExportProvider>();
                    services.AddTransient<IExcelImportProvider, ExcelImportProvider>();
                    services.AddTransient<IExcelExportProvider, ExcelExportProvider>();
                    break;
                default:
                    throw new NotImplementedException();
            }
            services.AddTransient<IWordExportService, WordExportService>();
            services.AddTransient<IExcelImportService, ExcelImportService>();
            services.AddTransient<IExcelExportService, ExcelExportService>();
            return services;
        }
    }
}
