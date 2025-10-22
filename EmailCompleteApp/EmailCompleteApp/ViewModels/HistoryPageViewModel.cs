using EmailCompleteApp.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailCompleteApp.ViewModels
{
    public partial class HistoryPageViewModel
    {
        private readonly SearchService _searchService;
        public HistoryPageViewModel()
        {
            _searchService = SearchService.Instance;
            //_ = InitializeHistoryData();
        }

    }
}

