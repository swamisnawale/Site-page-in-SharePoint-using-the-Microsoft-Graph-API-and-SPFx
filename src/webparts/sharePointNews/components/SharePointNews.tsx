import * as React from "react";
import styles from "./SharePointNews.module.scss";
import { ISharePointNewsProps, IState, NewsItem } from "./ISharePointNewsProps";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { TextField, PrimaryButton, DefaultButton } from "@fluentui/react";
import { addItemInList } from "./Services/Services";
import { __awaiter } from "tslib";

export default class SharePointNews extends React.Component<
  ISharePointNewsProps,
  IState
> {
  constructor(props: ISharePointNewsProps) {
    super(props);
    this.state = {
      AllNews: [],
      IsLoading: true,
      ShowForm: false,
      SiteID: "",
      Title: "",
      Description: "",
      Summary: "",
    };
  }

  componentDidMount(): void {
    this.getItems();
    this.getSharePointSiteID();
  }

  private getSharePointSiteID = () => {
    // https://graph.microsoft.com/v1.0/sites/{host-name}:/{server-relative-path}

    // https://learn.microsoft.com/en-us/graph/api/site-getbypath?view=graph-rest-1.0
    this.props.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3): void => {
        client
          .api(
            `https://graph.microsoft.com/v1.0/sites/${window.location.hostname}:/${this.props.context.pageContext.web.serverRelativeUrl}`
          )
          .version("v1.0")

          .get()
          .then((res: any) => {
            const siteID = res.id.split(",")[1];
            this.setState({
              SiteID: siteID,
            });
          });
      });
  };

  private getItems = () => {
    let topCount = 4999;
    let filterQuery = ``;
    let selectQuery = `ID,Title,Description,Summary,Created,PostLink`;
    let expandQuery = ``;
    let orderQuery = `ID desc`;
    let listName = "Latest News";
    let requestURL = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$top=${topCount}&$filter=${filterQuery}&$select=${selectQuery}&$expand=${expandQuery}&$orderby=${orderQuery}`;
    this.props.context.spHttpClient
      .get(requestURL, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json();
        }
      })
      .then((i) => {
        if (i.value.length == 0) {
          this.setState({
            IsLoading: false,
            AllNews: [],
          });
        } else {
          this.setState({
            IsLoading: false,
            AllNews: i.value,
          });
        }
      })
      .catch((err) => {
        this.setState({
          IsLoading: false,
          AllNews: [],
        });
      });
  };
  createPage = (newsTitle: string, newsDetail: string) => {
    const baseSitePage = {
      "@odata.type": "#microsoft.graph.sitePage",
      name: `${newsTitle}.aspx`,
      title: `${this.state.Title}`,
      pageLayout: "article",
      showComments: true,
      showRecommendedPages: false,
      titleArea: {
        enableGradientEffect: true,
        imageWebUrl:
          "https://cdn.hubblecontent.osi.office.net/m365content/publish/005292d6-9dcc-4fc5-b50b-b2d0383a411b/image.jpg",
        layout: "imageAndTitle",
        showAuthor: false,
        showPublishedDate: false,
        showTextBlockAboveTitle: false,
        textAboveTitle: "",
        textAlignment: "center",
        imageSourceType: 2,
        title: "sample1",
      },
      canvasLayout: {
        horizontalSections: [
          {
            layout: "oneColumn",
            id: "1",
            emphasis: "none",
            columns: [
              {
                id: "1",
                width: 12,
                webparts: [
                  {
                    id: "6f9230af-2a98-4952-b205-9ede4f9ef548",
                    innerHtml: `${newsDetail}`,
                  },
                ],
              },
            ],
          },
        ],
      },
    };

    this.props.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3): void => {
        client
          .api(
            `https://graph.microsoft.com/v1.0/sites/${this.state.SiteID}/pages`
          )
          .version("v1.0")

          .post(baseSitePage)
          .then((res: any) => {
            const pageID = res.id;

            console.log("pageID", pageID);
            this.props.context.msGraphClientFactory
              .getClient("3")
              .then((client: MSGraphClientV3): void => {
                client
                  .api(
                    `https://graph.microsoft.com/v1.0/sites/${this.state.SiteID}/pages/${pageID}/microsoft.graph.sitePage/publish`
                  )
                  .version("v1.0")
                  .post({})
                  .then((res) => {
                    console.log("Page published");

                    const json_AddProject: any = {
                      Title: this.state.Title,
                      Summary: this.state.Summary,
                      PostLink: {
                        Url: `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.state.Title}.aspx`,
                      },
                    };
                    addItemInList(
                      {
                        context: this.props.context,
                        listName: "Latest News",
                      },
                      json_AddProject
                    ).then((res) => {
                      this.setState({
                        Description: "",
                        Summary: "",
                        Title: "",
                        ShowForm: false,
                      });
                      console.log("News added in list");
                      this.getItems();
                    });
                  });
              });
          })
          .catch((err: any) => {
            console.log("Error", err);
          });
      });
  };

  private onTextChange = (newText: string) => {
    this.setState({
      Description: newText,
    });
    return newText;
  };

  addItem = () => {
    const promise = [this.createPage(this.state.Title, this.state.Description)];
    Promise.resolve(promise).then(() => {});
  };
  public render(): React.ReactElement<ISharePointNewsProps> {
    return (
      <div className={styles.sharePointNews}>
        <style>
          {`
          .ql-active .ql-toolbar{
            top:0 !important
          }
          `}
        </style>
        <div>
          <div className={styles.sharePointNews_FromContent}>
            <form
              action=""
              onSubmit={(e) => {
                e.preventDefault();
              }}
            >
              <TextField
                label="News title"
                onChange={(e, k: string) => {
                  this.setState({
                    Title: k,
                  });
                }}
                maxLength={255}
                value={this.state.Title}
              />
              <TextField
                label="News Summary"
                multiline
                onChange={(e, k: string) => {
                  this.setState({
                    Summary: k,
                  });
                }}
                maxLength={255}
                value={this.state.Summary}
              />
              <RichText
                value={this.state.Description}
                onChange={(text: string) => this.onTextChange(text)}
                label="News Details"
                className={styles.sharePointNews_RTE}
              />
              <div style={{ margin: "24px 0 0 0" }}>
                <PrimaryButton
                  text="Add news"
                  type="submit"
                  style={{ margin: "0 12px 0 0" }}
                  onClick={this.addItem}
                />
                <DefaultButton
                  text="Cancel"
                  type="reset"
                  onClick={() => [
                    this.setState({
                      Description: "",
                      Summary: "",
                      Title: "",
                      ShowForm: false,
                    }),
                  ]}
                />
              </div>
            </form>
          </div>
        </div>
        <div className={styles.sharePointNews_Content}>
          {this.state.AllNews.map((news: NewsItem) => {
            let dateStr = news.Created;
            let date = new Date(dateStr);

            // Define options for toLocaleDateString
            let options: {} = {
              year: "numeric",
              month: "short",
              day: "numeric",
            };

            // Format date
            let formattedDate = date.toLocaleDateString("en-US", options);

            return (
              <div
                className={styles.sharePointNews_NewsItem}
                onClick={() => {
                  window.open(news.PostLink.Url);
                }}
              >
                <img
                  src={require("../assets/default-image.png")}
                  className={styles.sharePointNews_NewsItem_Image}
                ></img>
                <div>
                  <p className={styles.sharePointNews_NewsItem_Title}>
                    {news.Title}
                  </p>
                  <p className={styles.sharePointNews_NewsItem_Description}>
                    {news.Summary}
                  </p>
                  <p>{formattedDate}</p>
                </div>
              </div>
            );
          })}
        </div>
      </div>
    );
  }
}
