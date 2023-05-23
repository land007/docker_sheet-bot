FROM land007/node:latest

MAINTAINER Yiqiu Jia <yiqiujia@hotmail.com>

#RUN . $HOME/.nvm/nvm.sh && cd / && npm install wxcrypt

RUN echo $(date "+%Y-%m-%d_%H:%M:%S") >> /.image_times && \
	echo $(date "+%Y-%m-%d_%H:%M:%S") > /.image_time && \
	echo "land007/sheet-bot" >> /.image_names && \
	echo "land007/sheet-bot" > /.image_name

COPY node /node_

ENV your_corp_id=\
	your_corp_secret=\
	doc_id=\
	share_url=\
	bot_key=


#docker build -t land007/sheet-bot:latest .
#> docker buildx build --platform linux/amd64,linux/arm64/v8,linux/arm/v7 -t land007/sheet-bot --push .
