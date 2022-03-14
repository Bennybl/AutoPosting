using System;
using System.Collections.Generic;
using System.Text;

namespace AutoPosting
{
    class ConfessionPost
    {
        public string Post {get;}

        public int PostNumber { get; }

        public DateTime DateCreated { get; }

        public bool IsPosted { get; set; }

        public Exception exceptionInPosting { get; set; }

        public ConfessionPost(string i_Post, int i_PostNumber)
        {
            Post = i_Post;
            PostNumber = i_PostNumber;
            DateCreated = DateTime.Now;
            IsPosted = true;
        }

        public override string ToString()
        {
            return "#" + PostNumber + "\n" + Post;
        }
    }
}
