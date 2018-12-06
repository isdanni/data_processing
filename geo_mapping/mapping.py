from geopy.geocoders import Nominatim
import pinyin
from cjklib.characterlookup import CharacterLookup

print ('Parsing start:')

# Convert Chinese location to Pinyin
print('Converting Chinese to Pinyin...')
print pinyin.get('北京', format="strip", delimiter="")
print('Converting finished.')

geolocator = Nominatim(user_agent="specify_your_app_name_here")
location = geolocator.geocode('ShenYang')
print(location.address)
print((location.latitude, location.longitude))
print(location.raw)
