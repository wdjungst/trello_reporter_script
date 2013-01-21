#!/usr/bin/env ruby
require 'rubygems'
require 'yaml'
require 'google_drive'
require 'pry'
require 'trello'
require 'optparse'
require 'mail'
require 'colorize'

include Trello
include Trello::Authorization

CONFIG = YAML::load_file(File.expand_path(File.dirname(__FILE__) + '/config/config.yml'))
TRELLO_REPORT = CONFIG['google_drive_key']
FILE = "./resources/sprint_number.txt"
ARCHIVE_FILE = "./resources/recent_archive_list.txt"
EMAIL_FILE = "./resources/msg.txt"

Trello::Authorization.const_set :AuthPolicy, OAuthPolicy
OAuthPolicy.consumer_credential = OAuthCredential.new(CONFIG['public_key'], CONFIG['secret'])
OAuthPolicy.token = OAuthCredential.new(CONFIG['token'], nil)

BOARD_ID = CONFIG['trello_board']

#List IDs - to get list id's run with option -i
CODE = CONFIG['list1']
VALIDATION = CONFIG['list2']
PROMOTE = CONFIG['list3']
REGRESSION = CONFIG['list4']
QA_MAINT = CONFIG['list5']
OTHER = CONFIG['list6']

@cards = []
@col = 5

def check_config
  if CODE.to_s == ''
    puts
    puts "*****You have not set up list id's in config.yml yet list ID's are listed below****".green
    list_ids
  end
end

def session
  session = GoogleDrive.login(CONFIG['user'], CONFIG['pass'])
  session.spreadsheet_by_key(TRELLO_REPORT)
end

def headers
  header = []
  template = session.worksheets.first
  for col in 1..template.num_cols
    header << template[1, col]
  end
  @col = header.count
  header
end

def user_input
  gets.to_i
end

def menu
  while true
    puts "1.  Add all items to board for Sprint #{File.open(FILE, &:readline)}"
    puts "2.  Change sprint number"
    puts "3.  Add additional items to sprint"
    puts "4.  Revert last archive"
    #puts "5.  Weekly update Done items to supervisor"
    puts "5.  Exit"

    choice = user_input
    case choice
      when 1
        update_sprint_sheet
      when 2
        change_sprint_number
      when 3
        add_items_to_sprint
      when 4
        revert_last_archive
      #when 5
      #  email_management
      when 5
        exit 0
      else
        puts "Invalid choice try again"
    end
  end
end

def validate_drive(sprint_num)
  ws = session.worksheets.last
  if ws.title != "Sprint #{sprint_num}"
    puts "Failed to create worksheet"
    exit 0
  end
end

#See comment below on supervisor_email method
#def email_management
#  @cards.clear
#  all_cards
#  supervisor_email
#end

def add_item(worksheet, card)
  users = []
  card.members.each { |member| users << member.full_name }
  last = (worksheet.num_rows + 1)
  worksheet[last, 1] = card.name
  worksheet[last, 2] = list_name(card)
  worksheet[last, 3] = users.join("\n")
  worksheet[last, 4] = card.description.strip!
  worksheet[last, 5] = card.short_id
  card.close!
  worksheet.save
end

def add_items_to_sprint
  sprint = File.open(FILE, &:readline).to_i
  sprint -= 1
  puts "Enter sprint number (latest update Sprint #{sprint}):"
  sprint_num = gets.strip!.to_i
  ss = session
  worksheets = ss.worksheets
  current = worksheets.first
  found = false
  worksheets.each_with_index do |ws, i|
    if ws.title == "Sprint #{sprint_num}"
      current = worksheets[i]
      found = true
    end
  end

  if found
    all_cards
    puts
    @cards.each_with_index do |c, i|
      i += 1
      puts "#{i}: #{c.name}"
    end
    print "card#: "
    choice = gets.to_i
    if choice > 0 && choice < @cards.count + 1
      puts @cards[choice - 1].name
      add_item(current, @cards[choice - 1])
    else
      puts "Invalid choice"
      @cards.clear
      menu
    end
    @cards.clear
  else
    puts
    puts "Sprint #{sprint_num} does not exist in list"
    puts
    menu
  end
end

def decrement_sprint
  num = File.open(FILE, &:readline)
  num = num.to_i
  num -= 1
  File.open(FILE, 'w') { |f| f.write(num.to_s) }
end

def validate_trello
  dec = false
  ws = session.worksheets.last
  items = ws.num_rows - 1

  if ws[2, 1] == ""
    dec = true
  end

  if @cards.count != items
    dec = true
    revert_last_archive
  end

  if dec
    decrement_sprint
    "Failed to add cards to worksheet"
    exit 0
  end
end

def revert_last_archive
  archived_cards = File.open(ARCHIVE_FILE, &:readline)
  archived_cards = archived_cards.split(",")
  archived_cards.each do |c|
    card = Card.find(c)
    card.closed = false
    card.update!
  end

end

def print_header
  ws = session.worksheets.last
  header = headers
  header.each_with_index do |h, i|
    ws[1, i + 1] = h
  end
  ws.save
end

def sprint_number
  number = File.open(FILE, &:readline)
  add_sprint_sheet(number)
  num = number.to_i
  num += 1
  File.open(FILE, 'w') { |f| f.write(num.to_s) }
end

def change_sprint_number
  puts "Enter current sprint number: "
  File.open(FILE, 'w') { |f| f.write(gets) }
end

def add_sprint_sheet(sprint_number)
  session.add_worksheet("Sprint #{sprint_number}", 100, @col)
  validate_drive(sprint_number)
end

def all_cards
  code = List.find(CODE)
  code_cards = code.cards
  code_cards.each do |c|
    @cards << c
  end

  validation = List.find(VALIDATION)
  validation_cards = validation.cards
  validation_cards.each do |c|
    @cards << c
  end

  promote = List.find(PROMOTE)
  promote_cards = promote.cards
  promote_cards.each do |c|
    @cards << c
  end

  regression = List.find(REGRESSION)
  regression_cards = regression.cards
  regression_cards.each do |c|
    @cards << c
  end

  maint = List.find(QA_MAINT)
  maint_cards = maint.cards
  maint_cards.each do |c|
    @cards << c
  end

  other = List.find(OTHER)
  other_cards = other.cards
  other_cards.each do |c|
    @cards << c
  end
end

def list_name(card)
  lid = card.list_id
  list = List.find(lid)
  list.name
end

def add_cards_to_sprint
  File.open(ARCHIVE_FILE, 'w') { |f| f.print "" }
  ws = session.worksheets.last
  @cards.each_with_index do |c, i|
    users = []
    c.members.each { |member| users << member.full_name }
    row = i + 2
    ws[row, 1] = c.name
    ws[row, 2] = list_name(c)
    ws[row, 3] = users.join("\n")
    ws[row, 4] = c.description.strip!
    ws[row, 5] = c.short_id
    if i == 0
      File.open(ARCHIVE_FILE, 'a') { |f| f.print "#{c.id}" }
    else
      File.open(ARCHIVE_FILE, 'a') { |f| f.print ",#{c.id}" }
    end
    c.close!
  end
  ws.save
  validate_trello
  notify_changes
end

#Nice to have but need to find a better way of truncating first week
#def supervisor_email
#  week = Date.today
#  File.open(EMAIL_FILE, 'w') do |f|
#    f.puts "Items done for the Week ending on #{week}"
#    f.puts
#    @cards.each_with_index do |c, i|
#      f.puts "#{i + 1}: #{c.name}"
#    end
#  end
#
#  f = ''
#  @cards.each_with_index do |c, i|
#    f = f + "#{i + 1}: #{c.name}\n"
#  end
#
#  mail = Mail.new do
#    from "#{CONFIG['user']}"
#    to "#{CONFIG['manager_email']}"
#    subject "Prof services QA completed items week ending on #{week}"
#    body "#{f}"
#    add_file :filename => "Week #{week.to_s.gsub(",", "")}.txt", :content => File.read(EMAIL_FILE)
#  end
#
#  mail.delivery_method :sendmail
#  mail.deliver
#end

def fill_body(array)
  b = ''
  array.each_with_index { |c, i| b = b + "#{i + 1}: #{c.name}\n"}
  b
end

def notify_changes
  sprint_num = File.open(FILE, &:readline)
  sprint_num = sprint_num.to_i
  sprint_num -= 1
  File.open(EMAIL_FILE, 'w') do |f|
    f.puts "Professional Services QA items completed Sprint #{sprint_num}"
    f.puts
    @cards.each_with_index do |c, i|
      f.puts "#{i + 1}: #{c.name}"
    end
  end

  b = ''
  @cards.each_with_index do |c, i|
    b = b + "#{i + 1}: #{c.name}\n"
  end

  mail = Mail.new do
    from "#{CONFIG['user']}"
    to "#{CONFIG['receipients']}"
    subject "Professional services QA completed items for Sprint #{sprint_num}"
    body b
    add_file :filename => "Sprint #{sprint_num}.txt", :content => File.read(EMAIL_FILE)
  end

  mail.delivery_method :sendmail
  mail.deliver
end

def update_sprint_sheet
  all_cards
  sprint_number
  print_header
  add_cards_to_sprint
end

def list_ids
  board = Board.find(BOARD_ID)
  lists = board.lists
  puts
  lists.each do |l|
    puts l.name
    puts l.id
    puts
  end
  puts "It is recommended that you update the ID's in the code and refactor list names then restart".green
  puts
  menu
end

def options
  options = {}
  optparse = OptionParser.new do |opts|
    options[:simple] = false
    options[:id] = false

    opts.on('-s', '--simple', 'Simple mode no menu') do
      options[:simple] = true
    end
    opts.on('-i', '--id', 'Get list id\'s') do
      options[:id] = true
    end
    opts.on('-h', '--help', 'Display this screen') do
      puts opts
      exit
    end
  end

  optparse.parse!

  if options[:simple]
    update_sprint_sheet
  elsif options[:id]
    list_ids
  else
    while true
      menu
    end
  end
end

check_config
options