# nominatim-excel
Adds [search](https://nominatim.org/release-docs/develop/api/Search/) and [reverse](https://nominatim.org/release-docs/develop/api/Reverse/) endpoints from [nominatim](https://nominatim.org/) [API](https://nominatim.org/release-docs/develop/api/Overview/) to Excel
#### Geocode(adressToSearch As String)
	allows you to look up a location from a text description
#### ReverseGeocode(lat As String, lng As String)
	brings an address from a latitude and longitude

![From Excel](https://i.imgur.com/eQjrzZV.png)

This project uses [VBA-Web fork](https://github.com/zbyna/VBA-Web) which fixs [UTF8 support for url encoding only in Windows](https://github.com/zbyna/VBA-Web/commit/dc87d7751d1ba9336aebfeb6b86b7fc258749781) It means this project does not work in macOS.


