s: install
	bundle exec jekyll server -H 0.0.0.0 -P 4006

run:
	cargo run

install:
	bundle install

pub:
	git status
	sleep 5
	git add .
	git commit -am 'update'
	git push

port:
	sudo dnf -y install firewalld
	sudo systemctl restart firewalld
	sudo firewall-cmd --permanent --zone=public --add-port=4007/tcp
	sudo firewall-cmd --reload

clean:
	cargo clean
	rm -rf _posts/* _site

pub2: clean
	cargo run
	make pub
