# coding: utf-8
require 'rubygems'
require 'net/imap'
require 'open-uri'
require 'fileutils'
require 'tmail'
require 'mail'
require 'sqlite3'

require 'tmail_mail_extension'

class String
  def is_binary_data?
    ( self.count( "^ -~", "^\r\n" ).fdiv(self.size) > 0.3 || self.index( "\x00" ) ) unless empty?
  end
end

class ParseEmail

  CTYPE_TO_EXT = {
      'image/jpeg' => 'jpg',
      'image/gif'  => 'gif',
      'image/png'  => 'png',
      'image/tiff' => 'tif'
  }

  def initialize
    @imap_server = Net::IMAP.new("imap.163.com")
    @imap_server.login('*****', '******')

    begin
      @test_db = SQLite3::Database.open "test.db"
      @test_db.execute "CREATE TABLE IF NOT EXISTS EmailInfos(FromAddr TEXT, ToAddr TEXT, BodyPlain TEXT, BodyHtml TEXT, ReceivedDate TIME)"
    rescue SQLite3::Exception => e
      puts "Exception occured"
      puts e
    end
  end

  def parse_email_by_tmail
    value_col = []
    @imap_server.select('INBOX')
    @imap_server.search(["ALL"]).each do |msg_id|
      whole_body = @imap_server.fetch(msg_id, 'RFC822')[0].attr['RFC822']
      mail_item = TMail::Mail.parse(whole_body)
      value_str = %Q( ('#{mail_item.from}', '#{mail_item.to}', '#{mail_item.body_plain}', '#{mail_item.body_html}', '#{mail_item.date.strftime("%Y-%m-%d %H:%M:%S")}') )
      value_col << value_str

      if mail_item.has_attachments?
        mail_item.parts.each_with_index do |part, index|
          file_dir = "attachment/#{mail_item.from}/"
          FileUtils.mkdir_p file_dir unless File.exist?(file_dir)
          file_name = (part['content-location'] && ['content-location'].body) || part.sub_header("content-type", "name") || part.sub_header("content-disposition", "filename")
          file_name ||= "#{index}.#{ext(part)}"
          file_name = "#{mail_item.date.strftime("%Y-%m-%d %H:%M:%S")}-" + file_name
          File.open(file_dir + file_name, "w+") do |f|
            f.write(part.body)
          end
        end
      end
      @imap_server.store(msg_id, "+FLAGS", [:SEEN])
    end

    if value_col.count > 0
      sql_str = <<-SQL
        INSERT INTO EmailInfos
        VALUES
        #{value_col.join(",")}
      SQL

      begin
        @test_db.execute sql_str
      rescue SQLite3::Exception => e
        puts "Exception occured."
        puts e
      ensure
        @test_db.close if @test_db
      end
    end
  end

  def ext(mail)
    CTYPE_TO_EXT[mail.content_type] || 'txt'
  end

  def self.run
    self.new.parse_email_by_tmail
  end
end

ParseEmail.run