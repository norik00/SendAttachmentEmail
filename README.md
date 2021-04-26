# SendAttachmentEmail
Send attachment email with use command line.

## Installing
```
$ pip intall git+https://github.com/norik00/SendAttachmentEmail
```

## Useage
Copy `config.sample.yml` and Rename to `config.yml`.  

**Setting of mail server**
`server`, `port` and `source`. `source` is sender email address.

**Setting of send destination**
`stores`, `branches` and `others`. Only `others` is cc.

example of Japanese
```yaml
stores:
    大阪店:
        branch: 関西支店,
        addressee: 
            - 大阪店　鈴木　様
            - 大阪店　木村　様
        mail:
            - oosaka-suzuki@sample.com
            - oosaka-kimura@sample.com
branches:
    関西支店:
        addressee:
            - 関西支店 佐藤　様
        mail:
            - kansai-sato@sample.com
        order:
others:
    関西支店:
        addressee:
            - 広島　様
        mail:
            - hiroshima@sample.com
```

**Setting of mail body**

`body` is mail body. This body use Python string formatting, So you can refer to your variable substitutions by name and use them in any order you want.
Variable is `addressee` and `division_ja`.

**Attachment directory**

`files` directory is same hierarchy directory where excute command.　When you want to change directory name, change `directory` at `config.yml`.
