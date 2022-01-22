#!/usr/bin/ruby
require 'uri'
require 'net/http'
require 'openssl'
require 'pry'
require 'json'
require 'google/apis/sheets_v4'
require 'googleauth'
require 'googleauth/stores/file_token_store'
require 'fileutils'
require 'yaml'

CFG_FILE = "fleetimporter.config"

cfg = YAML.load(File.open(CFG_FILE).read)
SAMSARA_AUTH_TOKEN = cfg["samsara_auth_token"]
SHEET_ID = cfg["sheet_id"]
SPECIAL_SHEET_NAME = cfg["special_sheet_name"]

if SHEET_ID.nil? || SPECIAL_SHEET_NAME.nil? || SAMSARA_AUTH_TOKEN.nil?
  puts "All configuration options must be present."
  exit 1
end

def get_all_equipment
  url = URI("https://api.samsara.com/fleet/equipment/stats?types=gpsOdometerMeters,gatewayEngineSeconds")

  http = Net::HTTP.new(url.host, url.port)
  http.use_ssl = true

  request = Net::HTTP::Get.new(url)
  request["Accept"] = 'application/json'
  request["Authorization"] = 'Bearer ' + SAMSARA_AUTH_TOKEN

  response = http.request(request)
  obj = JSON.parse(response.read_body)
  data = obj["data"]

  equipment = []
  data.each do |e|
	kms = (e["gpsOdometerMeters"]["value"].to_f / 1000).round(1)
	hours = (e["gatewayEngineSeconds"]["value"].to_f / 60 / 60).round(1)
	name = e["name"]
	#puts "#{name}: #{hours} hours, #{kms} km."
	equipment << OpenStruct.new({name:name, hours:hours, kms:kms})
  end
  return equipment
end

OOB_URI = "urn:ietf:wg:oauth:2.0:oob".freeze
APPLICATION_NAME = "Samsara Google Sheets integration".freeze
CREDENTIALS_PATH = "credentials.json".freeze
# The file token.yaml stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
TOKEN_PATH = "token.yaml".freeze
SCOPE = Google::Apis::SheetsV4::AUTH_SPREADSHEETS

##
# Ensure valid credentials, either by restoring from the saved credentials
# files or intitiating an OAuth2 authorization. If authorization is required,
# the user's default browser will be launched to approve the request.
#
# @return [Google::Auth::UserRefreshCredentials] OAuth2 credentials
def authorize_googleapi
  client_id = Google::Auth::ClientId.from_file CREDENTIALS_PATH
  token_store = Google::Auth::Stores::FileTokenStore.new file: TOKEN_PATH
  authorizer = Google::Auth::UserAuthorizer.new client_id, SCOPE, token_store
  user_id = "default"
  credentials = authorizer.get_credentials user_id
  if credentials.nil?
	url = authorizer.get_authorization_url base_url: OOB_URI
	puts "Open the following URL in the browser and enter the " \
		 "resulting code after authorization:\n" + url
	code = gets
	credentials = authorizer.get_and_store_credentials_from_code(
	  user_id: user_id, code: code, base_url: OOB_URI
	)
  end
  credentials
end

def clear_spreadsheet(service)
  service.clear_values(SHEET_ID, SPECIAL_SHEET_NAME)
end

def update_spreadsheet(service, newvalues)
  newdata = Google::Apis::SheetsV4::ValueRange.new
  newdata.major_dimension="ROWS"
  newdata.values = newvalues
  newdata.range = SPECIAL_SHEET_NAME

  updreq = Google::Apis::SheetsV4::BatchUpdateValuesRequest.new
  updreq.data = [newdata]
  updreq.value_input_option = "USER_ENTERED"

  response = service.batch_update_values(SHEET_ID, updreq)
end

# Initialize the API
service = Google::Apis::SheetsV4::SheetsService.new
service.client_options.application_name = APPLICATION_NAME
service.authorization = authorize_googleapi

all = get_all_equipment
new_cells = all.map {|x| [x.name, x.hours, x.kms] }
clear_spreadsheet(service)
update_spreadsheet(service, new_cells)
