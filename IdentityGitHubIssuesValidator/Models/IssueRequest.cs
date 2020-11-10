using System;
using System.Collections.Generic;
using System.Text;

namespace IdentityGitHubIssuesValidator.Models
{
    class IssueRequest
    {
        public string Owner { get; set; }

        public string RepoName { get; set; }

        public int IssueId { get; set; }
    }
}
