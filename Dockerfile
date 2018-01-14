FROM ruby:2.4
WORKDIR /usr/src/app
COPY Gemfile* ./
RUN bundle install
COPY . .

CMD ["./insanity-uploader.rb","--config","secrets-prodhost.yml","--par-report","pars-to-add.yml","-m","--sb"]
