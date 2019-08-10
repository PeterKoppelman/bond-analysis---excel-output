import auth_token as at

# emailfrom and password were set up by Peter Koppelman.
# They should still be active.
emailfrom = "python.automated.email.do.not.reply@gmail.com"
# put email addresses in the distribution list. They need to be in quotes.
distribution_list = ['']
emailto = "; ".join(distribution_list)
filetosend = at.file_name
username = emailfrom
password = "pythonnyu"
