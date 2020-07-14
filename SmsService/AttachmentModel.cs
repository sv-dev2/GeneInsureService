using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsuranceClaim.Models
{
    public class AttachmentModel
    {
        public Stream Attachment { get; set; }
        public string Name { get; set; }
    }
}
