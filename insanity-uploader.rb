#!/usr/bin/env ruby
####
# Copyright 2016-2017 John Messenger
# 
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
# 
# http://www.apache.org/licenses/LICENSE-2.0
# 
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
####

require 'rubygems'
require 'date'
require 'json'
require 'logger'
require 'mechanize'
require 'nokogiri'
require 'open-uri'
require 'rest-client'
require 'rubyXL'
require 'slop'
require 'yaml'


class String
  def casecmp?(other)
    self.casecmp(other).zero?
  end
end

####
# Log in to the Maintenance Database API
####
def login(api, username, password)
  login_request = {}
  login_request['user'] = {}
  login_request['user']['email'] = username
  login_request['user']['password'] = password

  begin
    res = api['users/sign_in'].post login_request.to_json, {content_type: :json, accept: :json}
  rescue RestClient::ExceptionWithResponse => e
    abort "Could not log in: #{e.response.to_s}"
  end
  res
end

####
# Search the Database for a task group with the specified name and return the parsed item
####
def find_task_group(api, name)
  search_result = api['task_groups'].get accept: :json, params: {search: name}
  tgs = JSON.parse(search_result.body)
  return nil if tgs.empty?
  thisid = tgs[0]['id']
  JSON.parse(api["task_groups/#{thisid}"].get accept: :json)
end

####
# Fetch the list of task groups from the database and return the parsed list
####
def find_task_groups(api)
  search_result = api['task_groups'].get accept: :json
  tgs = JSON.parse(search_result.body)
  return nil if tgs.empty?
  tgs
end

####
# Find a project.  If a task_group is specified, look there.  Otherwise, look at all projects.
####
def find_project_in_tg(api, tg, designation)
  if tg.nil?
    search_result = api["projects"].get accept: :json, params: {search: designation}
  else
    tgid = tg['id']
    search_result = api["task_groups/#{tgid}/projects"].get accept: :json, params: {search: designation}
  end
  projects = JSON.parse(search_result.body)
  return nil if projects.empty?
  thisid = nil
  projects.each do |project|
    thisdesig = project['designation']
    revdesig = designation + '-REV'
    if designation.casecmp?(thisdesig) || revdesig.casecmp?(thisdesig)
      thisid = project['id']
    end
  end
  return nil unless thisid
  JSON.parse(api["projects/#{thisid}"].get accept: :json)
end

####
# Create a new project in an existing Task Group
####
def add_project_to_tg(api, cookie, tg, newproj)
  tgid = tg['id']
  option_hash = {content_type: :json, accept: :json, cookies: cookie}
  begin
    res = api["task_groups/#{tgid}/projects"].post newproj.to_json, option_hash
    twit = 11
  rescue => e
    $logger.error "add_project_to_tg => exception #{e.class.name} : #{e.message}"
    if (ej = JSON.parse(e.response)) && (eje = ej['errors'])
      eje.each do |k, v|
        $logger.error "#{k}: #{v.first}"
      end
      exit(1)
    end
  end
  JSON.parse(res)
end

####
# Update a project that already exists
####
def update_project(api, cookie, proj, update)
  tgid = proj['task_group_id']
  projid = proj['id']
  option_hash = {content_type: :json, accept: :json, cookies: cookie}
  begin
    res = api["task_groups/#{tgid}/projects/#{projid}"].patch update.to_json, option_hash
    twit = 11
  rescue => e
    $logger.error "update_project => exception #{e.class.name} : #{e.message}"
    if (ej = JSON.parse(e.response)) && (eje = ej['errors'])
      eje.each do |k, v|
        $logger.error "#{k}: #{v.first}"
      end
      exit(1)
    end
  end
end

####
# Find an Event in a Project.
####
# @param [RestClient::Resource] api
# @param [proj_hash] proj
# @param [string] event_name
def find_event_in_proj(api, proj, event_name)
  projid = proj['id']
  tgid = proj['task_group_id']
  JSON.parse(api["task_groups/#{tgid}/projects/#{projid}/events"].get accept: :json, params: {search: event_name})
end

####
# Add events to a Project.  If an event of that name is already present, then update it.
####
# @param [RestClient::Resource] api
# @param [cookie] cookie
# @param [proj_hash] proj
# @param [Array] events
def add_events_to_project(api, cookie, proj, events)
  projid = proj['id']
  tgid = proj['task_group_id']
  option_hash = {content_type: :json, accept: :json, cookies: cookie}
  res = nil
  begin
    events.each do |event|
      ev = find_event_in_proj(api, proj, event[:name])
      if ev.empty?
        $logger.info("Adding event #{event[:name]} to project #{proj['designation']}")
        res = api["task_groups/#{tgid}/projects/#{projid}/events"].post event.to_json, option_hash
      else
        $logger.info("Patching event #{event[:name]} to project #{proj['designation']}")
        res = api["task_groups/#{tgid}/projects/#{projid}/events/#{ev.first['id']}"].patch event.to_json, option_hash
      end
    end
    twit = 11 ##########
  rescue => e
    $logger.error "add_events_to_project => exception #{e.class.name} : #{e.message}"
    if (ej = JSON.parse(e.response)) && (eje = ej['errors'])
      eje.each do |k, v|
        $logger.error "#{k}: #{v.first}"
      end
      exit(1)
    end
  end
end

###
# Parse the status message from the spreadsheet into a status value and an optional event
###
def parse_status(sts_string)
  s, e = sts_string.split(' - ')
  orig_s = s
  case
    when s.match(/WG\s*[bB]allot[- ]*[rR]ecirc/)
      s = 'WgBallotRecirc'
    when s.match(/WG\s*[bB]allot$/)
      s = 'WgBallot'
    when s.match(/TG\s*[bB]allot[- ]*[rR]ecirc/)
      s = 'TgBallotRecirc'
    when s.match(/TG\s*[bB]allot$/)
      s = 'TgBallot'
    when s.match(/Editor/)
      s = 'EditorsDraft'
    when s.match(/Sponsor\s*[bB]allot[- ]*[cC]ond/)
      s = 'SponsorBallotCond'
    when s.match(/Sponsor\s*[bB]allot$/)
      s = 'SponsorBallot'
    when s.match(/PAR\s*[dD]evelop/)
      s = 'ParDevelopment'
    when s.match(/PAR\s*[aA]pproved/)
      s = 'ParApproved'
  end
  events = []
  if e
    events << {date: Date.parse(e), name: s, description: orig_s + ': ' + Date.parse(e).to_s}
  end
  [s, events]
end

###
# Parse the Last Motion and Next Action from the spreadsheet into a standard form.
# If you don't give a value, you get 'Done'.
###
# @param [String] motion_string
def parse_motion(motion_string)
  s = motion_string
  case
    when s.nil? || s.empty?
      s = 'Done'
    when s.match(/WG\s*[bB]allot[- ]*[rR]ecirc/)
      s = 'WgBallotRecirc'
    when s.match(/WG\s*[bB]allot$/)
      s = 'WgBallot'
    when s.match(/TG\s*[bB]allot[- ]*[rR]ecirc/)
      s = 'TgBallotRecirc'
    when s.match(/TG\s*[bB]allot$/)
      s = 'TgBallot'
    when s.match(/Editor/)
      s = 'EditorsDraft'
    when s.match(/Sponsor\s*[bB]allot[- ]*[cC]ond/)
      s = 'SponsorBallotCond'
    when s.match(/Sponsor\s*[bB]allot$/)
      s = 'SponsorBallot'
    when s.match(/PAR\s*[dD]evelop/)
      s = 'ParDevelopment'
    when s.match(/PAR\s*[aA]pproval/)
      s = 'ParApproval'
    when s.match(/PAR\s*[mM]od/)
      s = 'ParMod'
    when s.match(/RevCom\s*[-*]\s*[cC]ond/)
      s = 'RevComCond'
    when s.match(/RevCom$/)
      s = 'RevCom'
    when s.match(/[wW]ithdraw/)
      s = 'Withdrawal'
  end
  s
end

###
# Parse the Last Motion and Next Action from the spreadsheet into a standard form.
# If you don't give a value, you get 'Done'.
###
# @param [String] desig_string
def parse_desig(desig_string)
  if (result = /P*(802(\.(\d+))*([A-Z]+|[a-z]+))([a-z]*)/.match(desig_string))
    base, unused, wg, projletters, amd = result.captures
    ptype = amd.empty? ? 'NewStandard' : 'Amendment'
  elsif (result = /P*(802(\.(\d+))([A-Z]+|[a-z]+))-[rR][eE][vV]/.match(desig_string))
    base, unused, wg, projletters = result.captures
    ptype = 'Revision'
  elsif (result = /P*(802(\.(\d+))([A-Z]+|[a-z]+)-*\d*)\/[cC][oO][rR]-*(\d+)/.match(desig_string))
    base, unused, wg, projletters, amd = result.captures
    ptype = 'Corrigendum'
  elsif (result = /P*(802(\.(\d+))([A-Z]+|[a-z]+)-*\d*)\/[eE][rR][rR]-*(\d+)/.match(desig_string))
    base, unused, wg, projletters, amd = result.captures
    ptype = 'Erratum'
  end

  [ptype, base]
end

###
# Parse the Last Motion and Next Action from the spreadsheet into a standard form.
# If you don't give a value, you get 'Done'.
###
# @param [String] short_date_string
def parse_short_date(short_date_string)
  if (result = /(jan|feb|mar|apr|may|june?|july?|aug|sep|oct|nov|dec)\s*(\d\d)/i.match(short_date_string))
    Date.parse($1 + " '" + $2)
  else
    Date.parse(short_date_string)
  end
end

####
# Delete a project from a task group, including its events
####
def delete_project(api, cookie, tg, project)
  option_hash = {accept: :json, cookies: cookie}
  events_result = api["task_groups/#{tg['id']}/projects/#{project['id']}/events"].get option_hash
  if events_result && !events_result.empty?
    events = JSON.parse(events_result)
    $logger.info "Project #{project['designation']} has #{events.count} events"
    events.each do |event|
      res = api["task_groups/#{tg['id']}/projects/#{project['id']}/events/#{event['id']}"].delete option_hash
    end
  end
  res = api["task_groups/#{tg['id']}/projects/#{project['id']}"].delete option_hash
  twit = 14
end

####
# Create a new item and add a request to it - UNUSED!
####
def add_new_item(api, cookie, number, subject, newreq)
  item = nil
  option_hash = {content_type: :json, accept: :json, cookies: cookie}
  newitem = {number: number, clause: newreq['clauseno'], date: newreq['date'], standard: newreq['standard'],
             subject: subject}
  res = api["items"].post newitem.to_json, option_hash
  if res.code == 201
    item = JSON.parse(res.body)
    reqres = add_request_to_item(api, cookie, item, newreq)
  end
  item
end

####
# Find a person
####
# @param [RestClient::Resource] api
# @param [string] first
# @param [string] last
# @param [string] role
####
def find_person(api, first, last, role)
  return nil if first.nil? || last.nil?
  res = api["people"].get accept: :json, params: { search: last }
  if res.code == 200
    people = JSON.parse(res.body)
    people.each do |pers|
      if pers['role'] == role && pers['first_name']&.casecmp?(first) && pers['last_name']&.casecmp?(last)
        return pers
      end
    end
  end
  nil
end

####
# Create a new person
####
# @param [RestClient::Resource] api
# @param [cookie] cookie
# @param [Hash] person
####
def add_new_person(api, cookie, person)
  pers = nil
  option_hash = {content_type: :json, accept: :json, cookies: cookie}
  res = api["people"].post person.to_json, option_hash
  if res.code == 201
    pers = JSON.parse(res.body)
  end
  pers
end

####
# Update an existing person
####
# @param [RestClient::Resource] api
# @param [cookie] cookie
# @param [Hash] perstoupdate
# @param [Hash] person
####
def update_person(api, cookie, perstoupdate, person)
  pers_id = perstoupdate['id']
  option_hash = {content_type: :json, accept: :json, cookies: cookie}
  pers = nil
  res = api["people/#{pers_id}"].patch person.to_json, option_hash
  if res.code == 201
    pers = JSON.parse(res.body)
  end
  pers
end

####
# Create a new task group
####
# @param [RestClient::Resource] api
# @param [cookie] cookie
# @param [string] abbrev
# @param [string] tgname
# @param [Hash] person
####
def add_new_task_group(api, cookie, abbrev, tgname, person)
  pers_id = person['id']
  newtg = { abbrev: abbrev, name: tgname, chair_id: pers_id }
  tg = nil
  option_hash = {content_type: :json, accept: :json, cookies: cookie}
  res = api["task_groups"].post newtg.to_json, option_hash
  if res.code == 201
    tg = JSON.parse(res.body)
  end
  tg
end

####
# Update an existing task group
####
# @param [RestClient::Resource] api
# @param [cookie] cookie
# @param [Hash] tgtoupdate
# @param [Hash] person
####
def update_task_group(api, cookie, tgtoupdate, person)
  tg_id = tgtoupdate['id']
  option_hash = {content_type: :json, accept: :json, cookies: cookie}
  tgtoupdate['chair_id'] = person['id']
  tg = nil
  res = api["task_groups/#{tg_id}"].patch tgtoupdate.to_json, option_hash
  if res.code == 201  # actually it seems to return 204.
    tg = JSON.parse(res.body)
  end
  tg
end

####
# Safely parse a date which might not be present
####
# @param [string] maybedate
def safe_date(maybedate)
  begin
    parsed_date = Date.parse(maybedate)
  rescue ArgumentError
    return nil
  end
  parsed_date
end

####
# Add or update projects from the Insanity Spreadsheet.  Optionally update the People and Task Groups
####
# @param [RestClient::Resource] api
# @param [cookie] cookie
# @param [string] filepath
# @param [Slop::Result] opts
def parse_insanity_spreadsheet(api, cookie, filepath, opts)
  book = RubyXL::Parser.parse filepath
  #book.worksheets.each do |w|
  #puts w.sheet_name
  #end

  if opts.people?
    peepsheet = book['People']
    people = []
    peepsheet[(i=0)..peepsheet.count-1].each do |peeprow|
      person = {
          role:        peeprow && peeprow[0].value,
          first_name:  peeprow && peeprow[1].value,
          last_name:   peeprow && peeprow[2].value,
          email:       peeprow && peeprow[3].value,
          affiliation: peeprow && peeprow[4].value
      }
      people << person
    end
    people.each do |person|
      pers = find_person(api, person[:first_name], person[:last_name], person[:role])
      if pers.nil?
        $logger.info("Adding new person #{person[:first_name]} #{person[:last_name]} as #{person[:role]}")
        add_new_person(api, cookie, person)
      else
        unless (pers['email'].casecmp?(person[:email])) && (pers['affiliation'].casecmp?(person[:affiliation]))
          $logger.info("Updating person #{person[:first_name]} #{person[:last_name]} as #{person[:role]}")
          update_person(api, cookie, pers, person)
        else
          $logger.debug("Up-to-date person #{pers['first_name']} #{pers['last_name']} as #{pers['role']}")
        end
       end
    end
  end

  tgsheet = book['TaskGroups']
  tgnames = {}
  tgsheet[(i=0)..tgsheet.count-1].each do |tgrow|
    tgnames[tgrow[0].value] = { name: tgrow[1].value, chair_first_name: tgrow[2]&.value, chair_last_name: tgrow[3]&.value }
  end
  tgnames.each do |abbrev, taskgroup|
    puts "TG #{abbrev}: #{taskgroup.to_s}" if $DEBUG
    if opts.task_groups? && ! /->/.match(abbrev)
      pers = find_person(api, taskgroup[:chair_first_name], taskgroup[:chair_last_name], 'Chair')
      if pers.nil?
        $logger.error("Chair #{taskgroup[:chair_first_name]} #{taskgroup[:chair_last_name]} not found" +
                          " for task group #{taskgroup[:name]}")
        next
      end
      tg = find_task_group(api, taskgroup[:name])
      if tg.nil?
        $logger.info "Creating task group #{taskgroup[:name]}."
        tg = add_new_task_group(api, cookie, abbrev, taskgroup[:name], pers)
      else
        $logger.debug "Updating existing task group #{taskgroup[:name]}"
        tg = update_task_group(api, cookie, tg, pers)
      end
    end
  end

  projsheet = book['Projects']
  projsheet[(i=1)..projsheet.count-1].each do |projrow|
    tgshortname = projrow && projrow[9]&.value
    tgname = tgnames[tgshortname][:name]
    tg = find_task_group(api, tgname)
    unless tg
      $logger.warn "Taskgroup #{tgname} not found"
      next
    end
    $logger.debug tgname
    desig = projrow && projrow[0]&.value
    if desig
      proj = find_project_in_tg(api, tg, desig)
      if proj.nil? || opts.update?
        $logger.info "Project #{desig} was not found for TG #{tgname}" if proj.nil?
        $logger.info "Project #{desig} will be updated" if opts.update?
        status, events = parse_status(projrow && projrow[3]&.value)
        ptype, base = parse_desig(desig)
        newproj = {
            designation: desig,
            project_type: ptype,
            base: base,
            short_title: projrow && projrow[1]&.value,
            title: 'unset',
            draft_no: projrow && projrow[4]&.value,
            status: status,
            last_motion: parse_motion(projrow && projrow[2]&.value),
            next_action: parse_motion(projrow && projrow[5]&.value),
            award: projrow && projrow[14]&.value
        }
        if opts.delete_existing? && !proj.nil?
          $logger.info "Deleting existing project #{desig}"
          delete_project(api, cookie, tg, proj)
        end
        $logger.info "Adding project #{desig} to TG #{tgname}"
        proj = add_project_to_tg(api, cookie, tg, newproj)
        unless proj
          $logger.error "Addition failed."
          raise('ProjAdditionFailed')
        end
        # then add extra stuff to it like events
        if projrow && projrow[6]&.value
          date = parse_short_date(projrow[6]&.value)
          events << {date: date, name: 'PAR ends', description: "PAR ends: #{date.to_s}"}
        end
        if projrow && projrow[11]&.value
          date = parse_short_date(projrow[11]&.value)
          events << {date: date, end_date: date + 213, name: 'Pool', description: "Sponsor ballot pool: #{date.to_s}"}
        end
        if projrow && projrow[12]&.value
          date = parse_short_date(projrow[12]&.value)
          events << {date: date, end_date: date + 30, name: 'MEC', description: "Manadatory Editorial Co-ordination: #{date.to_s}"}
        end

        unless events.empty?
          add_events_to_project(api, cookie, proj, events)
        end
      else
        $logger.info "Project #{desig} exists as #{proj['short_title']}"
      end

    else
      $logger.info "Skipping undesignated project in row #{i}"
    end
    #exit
  end
end

####
# Follow the PAR detail link to get the PAR's dates, full title, etc.
####
# @param [Mechanize] agent
# @param [uri] link
def parse_par_page(agent, link)
  events = []
  fulltitle = ''
  projpage = agent.get(link)
  box = projpage.css('div.tab-content-box')
  par_path = box.css('div.task_menu').children.first.attributes['href'].to_s
  par_url = URI.parse(link) + URI.parse(par_path)
  (0..box.children.count-1).each do |parlineno|
    ptype = ''
    case box.children[parlineno].to_s
      when /Type of Project/
        case box.children[parlineno+1].to_s
          when /Modify Existing/
            ptype = 'Modification'
          when /Revision to/
            ptype = 'Revision'
          when /Amendment to/
            ptype = 'Amendment'
          when /New IEEE/
            ptype = 'New'
        end
      when /PAR Request Date/
        mydate = safe_date(box.children[parlineno+1].to_s)
        if mydate
          name = if ptype == 'Modification'
                   'PAR Modification Requested'
                 else
                   'PAR Requested'
                 end
          events << {date: mydate, name: name, description: "#{name}: " + mydate.to_s}
        end
      when /PAR Approval Date/
        mydate = safe_date(box.children[parlineno+1].to_s)
        if mydate
          name = if ptype == 'Modification'
                   'PAR Modification Approval'
                 else
                   'PAR Approval'
                 end
          events << {date: mydate, name: name, description: "#{name}: " + mydate.to_s}
        end
      when /PAR Expiration Date/
        mydate = safe_date(box.children[parlineno+1].to_s)
        name = 'PAR Expiry'
        events << {date: mydate, name: name, description: "#{name}: " + mydate.to_s} if mydate
      when /2.1 Title/
        if box.children[parlineno].css('td.b_align_nw').empty?
          fulltitle = box.children[parlineno+1].to_s
        else
          fulltitle = box.children[parlineno].css('td.b_align_nw')[0].children[1].to_s
        end
      when /4.2.*Initial Sponsor Ballot/
        mydate = safe_date(box.children[parlineno+1].to_s)
        name = 'Expected Initial Sponsor Ballot'
        events << {date: mydate, name: name, description: "#{name}: " + mydate.to_s} if mydate
      when /4.3.*RevCom/
        mydate = safe_date(box.children[parlineno+1].to_s)
        name = 'Expected RevCom'
        events << {date: mydate, name: name, description: "#{name}: " + mydate.to_s} if mydate
    end
  end
  return fulltitle, par_url.to_s, events
end

####
# Add or update projects from the Development Server's Active PARs page.
####
# @param [RestClient::Resource] api
# @param [cookie] cookie
# @param [string] dev_host
# @param [string] user
# @param [string] pw
def update_projects_from_dev_server(api, cookie, dev_host, user, pw)
  agent = Mechanize.new
  if $DEBUG
    agent.set_proxy('localhost', 8888)
    agent.agent.http.verify_mode = OpenSSL::SSL::VERIFY_NONE
  end

  $logger.info("Updating projects from Development Server")
  # Assume that we are not logged in, and log in to the Development server
  page = agent.get('http://' + dev_host)
  f = page.forms.first
  f.x1 = user
  f.x2 = pw
  f.f0 = '3' # myproject
  page = agent.submit(f, f.buttons.first)
  #puts page.pretty_print_inspect if $DEBUG
  nextlink = URI::HTTP.build(host: dev_host, path: '/pub/active-pars', query: 's=802.1')
  # Find the "Active PARs" page.
  # Process each page of the list of active projects
  until nextlink.nil?
    $logger.debug("New page with nextlink #{nextlink}")
    searchresult = agent.get(nextlink)
    # puts searchresult.pretty_print_inspect
    # Examine each data row representing an active project.  Extract dates to create events in the project timeline.
    searchresult.parser.css('tr.b_data_row').each do |row|
      tds = row.css('td')
      desig = tds[1].children.first.children.to_s
      next unless /802.1[a-zA-Z]+|802[a-zA-Z]/.match(desig)
      $logger.debug("Considering project #{desig}")
      events = []
      par_link = tds[1].children.first.attributes['href'].to_s
      par_url = tds[3].children.first.children.to_s
      par_approval = safe_date(tds[4].children.css('noscript').children.to_s)
      events << { date: par_approval, name: 'PAR Approval', description: 'PAR Approval: ' + par_approval.to_s } if par_approval

      fulltitle,ign,e = parse_par_page(agent, par_link)
      events += e

      # Look up the project in the database without using a task group
      desig[/^P*/] = '' # Remove leading P
      proj = find_project_in_tg(api, nil, desig)
      if proj.nil?
        $logger.warn("Expected project #{desig} (from devserv) not found in database")
        next
      else
        $logger.debug("Matching PAR #{desig} to project #{proj['designation']}")
      end
      # Overwrite existing project information and add new events to the project.
      add_events_to_project(api, cookie, proj, events) unless events.empty?
      update_project(api, cookie, proj, { title: fulltitle, par_url: par_url }) unless fulltitle.empty? and
          par_url.empty?
      twit = 34
    end
    # Find the link to the next page of projects.
    pager = searchresult.parser.css('div.pager').children
    nextstr = pager[-1].children[-1].children.to_s
    nextlink = nextstr.empty? ? nil : pager.css('a')[-1].attributes['href'].to_s
  end

end

####
# Add or update projects from the Development Server's PAR report page.
# It wouldn't be a good idea to just add them all: There are multiple projects with the same designation.
# This is because revision projects don't have unique names.  Therefore, we use a list of names
# read from a file.
####
# @param [RestClient::Resource] api
# @param [cookie] cookie
# @param [string] dev_host
# @param [string] user
# @param [string] pw
# @param [Hash] projects
# @param [Array] task_groups
def update_projects_from_par_report(api, cookie, dev_host, user, pw, projects, task_groups)
  agent = Mechanize.new
  if $DEBUG
    agent.set_proxy('localhost', 8888)
    agent.agent.http.verify_mode = OpenSSL::SSL::VERIFY_NONE
  end

  $logger.info("Updating projects from PAR report on development server")
  # Assume that we are not logged in, and log in to the Development server
  page = agent.get('http://' + dev_host)
  f = page.forms.first
  f.x1 = user
  f.x2 = pw
  f.f0 = '3' # myproject
  page = agent.submit(f, f.buttons.first)
  #puts page.pretty_print_inspect if $DEBUG

  nextlink = URI::HTTP.build(host: dev_host, path: '/pub/par-report', query: 'par_report=1&committee_id=&s=802.1')

  # Find the "PAR Report" page.
  # Process each page of the list of active projects
  until nextlink.nil?
    $logger.debug("New PAR report page with nextlink #{nextlink}")
    searchresult = agent.get(nextlink)
    # puts searchresult.pretty_print_inspect
    # Examine each data row representing a project.  Extract dates to create events in the project timeline.
    searchresult.parser.css('tr.b_data_row').each do |row|
      tds = row.css('td')
      desig = tds[0].children.first.children.to_s
      next unless /802.1[a-zA-Z]+|802[a-zA-Z]/.match(desig)
      partype = tds[1].children.to_s
      $logger.debug("Considering project #{desig} (#{partype})")
      events = []
      par_link = tds[0].children.first.attributes['href'].to_s

      par_approval = safe_date(tds[6].children.to_s)
      events << { date: par_approval, name: 'PAR Approval', description: 'PAR Approval: ' + par_approval.to_s } if par_approval

      par_expiry = safe_date(tds[7].children.to_s)
      events << { date: par_expiry, name: 'PAR Expiry', description: 'PAR Expiry: ' + par_expiry.to_s } if par_expiry

      status = tds[10].children.to_s

      fulltitle,par_url,e = parse_par_page(agent, par_link)
      events += e

      # Look up the project in the database without using a task group
      desig[/^P*/] = '' # Remove leading P
      unless projects.keys.include? desig
        $logger.info("Not adding project #{desig} (#{partype}) (from PAR Report) as it's not in the approved list")
        next
      end

      proj = find_project_in_tg(api, nil, desig)
      if proj.nil?
        $logger.warn("Expected project #{desig} (from PAR report) not found in database: adding it")
        # Find the task group in the task_groups list from the projects[desig] entry
        tg = task_groups.detect { |t| t['abbrev'] == projects[desig]}
        unless tg
          $logger.error("Task group #{projects[desig]} from approved list not found")
          next
        end
        # Add the project to the task group.
        case status
          when 'Complete'
            status = 'Approved'
          when 'WG Draft Development'
            status = 'ParApproved'
          when 'Sponsor Ballot: Invitation'
            status = 'WgBallotRecirc'     # This is a bit of a kludge
          when /Sponsor Ballot/
            status = 'SponsorBallot'
          when /Com Agenda (\d\d-\w+-\d\d\d\d)/
            agd = safe_date(status)
            status = status.split[0]            # Just keep the first word: NesCom or RevCom
            events << { date: agd, name: status, description: "#{status}: " + agd.to_s } if agd
        end
        ptype, base = parse_desig(desig)
        newproj = {
            designation: desig,
            project_type: ptype,
            base: base,
            short_title: 'unset',
            title: 'unset',
            status: status,
            next_action: 'EditorsDraft'     # This is a kludge
        }
        proj = add_project_to_tg(api, cookie, tg, newproj)
      else
        $logger.debug("Matching PAR #{desig} to project #{proj['designation']}")
      end
      # Overwrite existing project information and add new events to the project.
      add_events_to_project(api, cookie, proj, events) unless events.empty?
      update_project(api, cookie, proj, { title: fulltitle, par_url: par_url }) unless fulltitle.empty? and
          par_url.empty?
      twit = 34
    end
    # Find the link to the next page of projects.
    pager = searchresult.parser.css('div.pager').children
    nextstr = pager[-1].children[-1].children.to_s
    nextlink = nextstr.empty? ? nil : pager.css('a')[-1].attributes['href'].to_s
  end

end

####
# Given the URL of a Ballot announcement, parse the text and return a hash containing the fields
####
# @param [string] url
# @param [Array] creds
####
def parse_announcement(url, creds)
  rqstream = open(url, http_basic_authentication: creds)
  rqdoc = Nokogiri::HTML(rqstream)
  text = rqdoc.xpath('//body//text()').to_s     # this is a really cool line
  return nil unless /^NOTE.*ALL.*RESPONSES/.match(text)
  return nil unless /^INCLUDE COMMENTS ONLY/.match(text)
  fields = {}
  # Get the date from the "Head-of-Message" because that's the date it was *really* posted.
  rqdoc.at('ul').search('li').each do |li|
    em = li.at('em')
    if /Date/.match(em.children.to_s)
      fields['date'] = Date.parse(li.children[1].to_s[2..-1])
      break
    end
  end

  # Get the remaining fields from the stylised form
  # NAME
  matches = /^TO:\s*(?<name>.+)\n/.match(text)
  unless matches
    $logger.warn "Parse error (TO) in Request #{url}"
    return nil
  end
  fields.merge!(Hash[ matches.names.zip(matches.captures)])

  # The other sections are parsed using a line-based scheme.  This isn't very
  # good.  What's more, it turns out that there can be embedded HTML in the message body.
  bin = :bin
  collection = {}
  textcollection = ''
  text.each_line do |line|
    if /^The 802.1 voting members that are entitled/i.match(line)
      collection[bin] = textcollection.strip
      textcollection = ''
      bin = :voters
      next
    elsif /^The closing date of this/i.match(line)
      collection[bin] = textcollection.strip
      textcollection = ''
      bin = :closing
      next
    elsif /can be found at/i.match(line)
      collection[bin] = textcollection.strip
      textcollection = ''
      bin = :bin
      next
    end
    break if /==========/.match(line) #and bin != :bin
    textcollection << line
  end
  collection[bin] = textcollection.strip
  fields['voters'] = collection[:voters]
  fields['closing'] = Date.parse(collection[:closing]) if collection[:closing]
  twit = 10
  fields
end

####
# Update projects from the Mail Server archive.
####
# @param [RestClient::Resource] api
# @param [cookie] cookie
# @param [string] arch_url
# @param [string] mailstart
# @param [string] user
# @param [string] pw
# @param [Array] blacklist
def update_projects_from_mail_server(api, cookie, arch_url, mailstart, user, pw, blacklist)
  #
  # Parse each index page of the 802.1 Maintenance Reflector Archive
  # and find the maintenance items
  #
  em_arch_url = arch_url + '/' + mailstart
  mtarch_creds = [user, pw]
  num_messages = 0
  num_responses = 0
  num_malformed_title = 0
  num_ballots = 0
  num_unparseable_announcements = 0
  num_events_added = 0

  $logger.info("Updating projects from Mail Archive")
  catch :done do
    while em_arch_url do
      page = open(em_arch_url, http_basic_authentication: mtarch_creds) { |f| f.read }
      pagedoc = Nokogiri::HTML(page)
      pagedoc.at('ul').search('li').each do |el|
        next unless el.children[0].name =~ /strong/
        # For each message...
        num_messages += 1
        href = el.children[0].children[0].attributes['href'].to_s
        url = arch_url + '/' + href
        titlestr = el.children[0].children[0].children[0].to_s
        if titlestr =~ /^[Rr][Ee]/ # discard responses to other maintenance items
          num_responses += 1
          next
        end
        mtchdata = /^\[802.1 - (?<number>\d+)\]\s+(?<type>\w+)\sgroup\s+(?<recirc>recirc\w*)?\s*ballot\s+of\s+(?<draft>P?802\S+)/i.match(titlestr)
        unless mtchdata
          num_malformed_title += 1
          next
        end
        num_ballots += 1
        number = '%d' % mtchdata['number']
        btype = mtchdata['type']
        recirc = mtchdata['recirc']
        draft = mtchdata['draft']
        $logger.debug "#{number}: #{btype} group #{recirc} ballot of #{draft}: #{url}"
        if blacklist&.include? number
          $logger.info "Ignoring blacklisted item #{number}"
          next
        end

        announcement = parse_announcement(url, mtarch_creds)
        if announcement.nil?
          $logger.warn("Couldn't parse ballot announcement #{url}")
          num_unparseable_announcements += 1
          next
        end
        desig,draftno = draft.split('/D')
        desig[/^P*/] = '' # Remove leading P
        proj = find_project_in_tg(api, nil, desig)
        if proj.nil?
          $logger.warn("Expected project #{desig} (from mailserv) not found in database: #{url}")
          next
        end
        startev = find_event_in_proj(api, proj, 'PAR Approval')
        if startev.empty?
          $logger.warn("Project #{desig} has no PAR Approval date")
          next
        end
        if announcement['date'] < Date.parse(startev&.first['date'])
          $logger.info("Not adding ballot on #{draft} as it starts before #{startev&.first['date'].to_s}")
          next
        end
        events = [ {
                       date: announcement['date'],
                       end_date: announcement['closing'],
                       name: "#{/task/i.match(btype) ? 'TG' : 'WG'} #{recirc ? 'recirc' : 'ballot'}: D#{draftno}",
                       description: "#{btype} group #{recirc ? 'recirculation ' : ''}ballot of #{draft}"
                   } ]
        add_events_to_project(api, cookie, proj, events)
        num_events_added += 1
        twit = 8
      end
      nextpage = pagedoc.search('tr')[1].children.search('td').children[4].attributes['href']
      em_arch_url = nextpage ? arch_url + '/' + nextpage.value : nil
    end
  end


  puts("num_requests: #{num_messages}\n")
  puts("num_responses: #{num_responses}\n")
  puts("num_malformed_title: #{num_malformed_title}\n")
  puts("num_unparseable_announcements: #{num_unparseable_announcements}\n")
  puts("num_events_added: #{num_events_added}\n")

rescue StopIteration

end

#
# Main program
#
begin
  opts = Slop.parse do |o|
    o.string '-s', '--secrets', 'secrets YAML file name', default: 'secrets.yml'
    o.bool   '-d', '--debug', 'debug mode'
    o.bool   '-x', '--delete-existing', 'delete existing projects before creating new ones'
    o.bool   '-u', '--update', 'update existing projects'
    o.bool   '-t', '--task-groups', 'Update the Task Groups from the Insanity spreadsheet'
    o.string '-f', '--filepath', 'path to an Excel XLSX-format Insanity spreadsheet'
    o.bool   '-b', '--devserv', 'update from development server'
    o.string '-r', '--par-report', 'add projects listed in file name from Devserv\'s PAR report'
    o.bool   '-m', '--mailserv', 'update from mailing list'
    o.bool   '-p', '--people', 'create or update people from the Insanity spreadsheet\'s People tab'
  end

  config = YAML.load(File.read(opts[:secrets]))
#
# Log in to the 802.1 Maintenance Database
#
  $DEBUG = opts.debug?
  $logger = Logger.new(STDOUT)
  $logger.level = Logger::INFO
  $logger.level = Logger::DEBUG if $DEBUG

  if $DEBUG
    RestClient.proxy = "http://localhost:8888"
    $logger.debug("Using HTTP proxy #{RestClient.proxy}")
  end

  maint = RestClient::Resource.new(config['api_uri'], :verify_ssl => ! $DEBUG)
  res = login(maint, config['email'], config['password'])
# Save the session cookie
  maint_cookie = {}
  res.cookies.each {|ck| maint_cookie[ck[0]] = ck[1] if /_session/.match(ck[0])}

  parse_insanity_spreadsheet(maint, maint_cookie, opts[:filepath], opts) if opts[:filepath]

  task_groups = find_task_groups(maint)

  if opts[:par_report]
    update_projects_from_par_report( maint, maint_cookie, config['dev_host'], config['dev_user'],
                                     config['dev_pw'], YAML.load(File.read(opts[:par_report])), task_groups)
  end
  if opts.devserv?
    update_projects_from_dev_server(maint, maint_cookie, config['dev_host'], config['dev_user'],
                                    config['dev_pw'])
  end

  if opts.mailserv?
    update_projects_from_mail_server(maint, maint_cookie, config['email_archive'],
                                     config['email_start'],
                                     config['archive_user'], config['archive_password'],
                                     config['blacklist'])
  end
end
