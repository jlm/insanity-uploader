FROM ruby:2.4
WORKDIR /usr/src/app
COPY Gemfile* ./
RUN bundle install
COPY . .

CMD ["./insanity-uploader.rb","--config","secrets-prodhost.yml","--par-report","http://ieee802.org/1/files/public/insanity/pars-to-add.yml","-m","--sb"]
