# jira-tool

#### 介绍
Jira data obtain from JQL

#### 软件架构

```
def search_issues(
                    jql_str: str,
                    startAt: int = 0,
                    maxResults: int = 50,
                    validate_query: bool = True,
                    fields: str | List[str] | None = "*all",
                    expand: str | None = None,
                    json_result: bool = False
) -> (Dict[str, Any] | ResultList[Issue])
```


#### 安装教程（依赖库）

1.  > pip install xlwt
2.  > pip install jira
3.  > pip install rich

#### 使用说明

usage: VizProject.py --help

********** Jira Data Visualization by Project **********


- optional arguments:
-   -h, --help            show this help message and exit
-   -v, --version         show program's version number and exit
-   --project-id PROJECT_ID [PROJECT_ID ...]
-                         (可选参数)通过project_id来搜索Jira数据, ex: X32A0-T972
-   --status STATUS [STATUS ...]
-                         (可选参数)通过status来搜索Jira数据, ex: OPEN, Resolved
-   --reporter REPORTER [REPORTER ...]
-                         (可选参数)通过reporter来搜索Jira数据, ex: san.zhang
-   --component COMPONENT [COMPONENT ...]
-                         (可选参数)通过component来搜索Jira数据, ex: HDMI, Dolby Vsion
-   --resolution RESOLUTION [RESOLUTION ...]
-                         (可选参数)通过resolution来搜索Jira数据, ex: Resolved, Won't fix
-   --priority PRIORITY [PRIORITY ...]
-                         (可选参数)通过priority来搜索Jira数据, ex: P0, P1
-   --severity SEVERITY [SEVERITY ...]
-                         (可选参数)通过severity来搜索Jira数据, ex: Normal, Major
-   --label LABEL [LABEL ...]
-                         (可选参数)通过label来搜索Jira数据, ex: pmlist-zql-20230103
-   --month MONTH [MONTH ...]
-                         (可选参数)通过date月份日期来搜索Jira数据, ex: 2022-11
-   --duration DURATION [DURATION ...]
-                         (可选参数)通过date月份日期范围来搜索Jira数据, ex: 2022-11 2023-02
-   --date-range DATE_RANGE [DATE_RANGE ...]
-                         (可选参数)通过date月份日期来筛选目标时间范围内的Jira数据内容, ex: 2022-11 2023-02
-   --testcase-id TESTCASE_ID [TESTCASE_ID ...]
-                         (可选参数)通过testcase_id来搜索Jira数据, ex: TV-F3081F0001
-   --testcase-check      (可选参数)加上该参数会进行testcase_id检测, 默认: False
-   --active-check        (可选参数)搜索统计Jira数据中所有人员comment活跃度占比, 默认: False
-   --label-check LABEL_CHECK [LABEL_CHECK ...]
-                         (可选参数)搜索统计Jira数据中添加labels人员的占比, ex: Common_From_Project, SH-Support-2023
-   --verify-check        (可选参数)搜索统计Jira数据中verified人员的占比
-   --di-count            (可选参数)搜索统计Severity并计算整体DI值
-   --raw-command JQL     (可选参数)通过JQL语句来搜索Jira数据, ex: "Project ID" = AM30A2-T950D4 AND status in (OPEN, Reopened)"
-   -e, --expand          (可选参数)搜索范围加入changelog的历史操作数据, 默认: False
-   -o, --output          (可选参数)保存数据到本地excel表格, 表格默认命名: Output_Result_YYYYMMDD_HHMMSS.xlsx, 默认: False
-   --verbose             (可选参数)加上该参数会打印更多调试信息, 默认: False

#### 参与贡献

1.  Fork 本仓库
2.  新建 Feat_xxx 分支
3.  提交代码
4.  新建 Pull Request


#### 特技

1.  使用 Readme\_XXX.md 来支持不同的语言，例如 Readme\_en.md, Readme\_zh.md
2.  Gitee 官方博客 [blog.gitee.com](https://blog.gitee.com)
3.  你可以 [https://gitee.com/explore](https://gitee.com/explore) 这个地址来了解 Gitee 上的优秀开源项目
4.  [GVP](https://gitee.com/gvp) 全称是 Gitee 最有价值开源项目，是综合评定出的优秀开源项目
5.  Gitee 官方提供的使用手册 [https://gitee.com/help](https://gitee.com/help)
6.  Gitee 封面人物是一档用来展示 Gitee 会员风采的栏目 [https://gitee.com/gitee-stars/](https://gitee.com/gitee-stars/)
