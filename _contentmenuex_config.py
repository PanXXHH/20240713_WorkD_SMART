from mysupport.WindowContentMenuEx.configer import asconfig
import os
import datetime

@asconfig
def 添加工作文件():

    sys_hosts = {"DEFAULT" : "SYS_DEFAULT"}

    hosts = {r"^@$": "PIN",
             r"^%$": "DEFINE",
             r"^%!$": "DEFINE_FORCE",
             r"^(%)\((\d+)\)$": "DEFINE_INDEX"
             }
             
    return dict(locals())

@asconfig
def folder_limit():
    limit = 10
    dirname = "历史档案"
    matchlist = [r"^\d{4}-\d{2}-\d{2}.txt$",
                 r"^\d{4}-\d{2}-\d{2}.yml$",
                 r"^\d{4}-\d{2}-\d{2}.mpp$"]

    return dict(locals())