using System.ComponentModel.DataAnnotations;

namespace DataProfile.Models
{
    public class Profile
    {
        public int Id { get; set; }

        [Required(ErrorMessage = "Nama Lengkap wajib diisi.")]
        [Display(Name = "Nama Lengkap")]
        public string FullName { get; set; } = string.Empty;

        [Required(ErrorMessage = "Email wajib diisi.")]
        [EmailAddress(ErrorMessage = "Format email tidak valid.")]
        [Display(Name = "Email")]
        public string Email { get; set; } = string.Empty;

        [Required(ErrorMessage = "Nomor Telepon wajib diisi.")]
        [Display(Name = "Nomor Telepon")]
        public string PhoneNumber { get; set; } = string.Empty;

        [Required(ErrorMessage = "Tanggal Lahir wajib diisi.")]
        [Display(Name = "Tanggal Lahir")]
        [DataType(DataType.Date)]
        public DateTime BirthDate { get; set; }

        [Required(ErrorMessage = "Alamat wajib diisi.")]
        [Display(Name = "Alamat")]
        public string Address { get; set; } = string.Empty;
    }
}