// export const getItems = () => {
//     let topCount = 4999;
//     let filterQuery = `IsActive eq '1'`;
//     let selectQuery = `*`;
//     let expandQuery = ``;
//     let orderQuery = `ID desc`;
//     let requestURL = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items?$top=${topCount}&$filter=${filterQuery}&$select=${selectQuery}&$expand=${expandQuery}&$orderby=${orderQuery}`;
//     this.props.context.spHttpClient
//       .get(requestURL, SPHttpClient.configurations.v1)
//       .then((response: SPHttpClientResponse) => {
//         if (response.ok) {
//           return response.json();
//         }
//       })
//       .then((i) => {
//         if (i.value.length == 0) {
//           this.setState({
//             IsLoading: false,
//             ListExist: true,
//             ListItems: [],
//           });
//         } else {
//           this.setState({
//             IsLoading: false,
//             ListItems: i.value,
//           });
//         }
//       })
//       .catch((err) => {
//         this.setState({
//           ListExist: false,
//           IsLoading: false,
//           ListItems: [],
//         });
//       });
//   };

//   export const createPage = () => {
//     const baseSitePage = {
//       "@odata.type": "#microsoft.graph.sitePage",
//       name: "News demo 3.aspx",
//       title: "News demo 3",
//       pageLayout: "article",
//       showComments: true,
//       showRecommendedPages: false,
//       titleArea: {
//         enableGradientEffect: true,
//         imageWebUrl:
//           "https://cdn.hubblecontent.osi.office.net/m365content/publish/005292d6-9dcc-4fc5-b50b-b2d0383a411b/image.jpg",
//         layout: "imageAndTitle",
//         showAuthor: false,
//         showPublishedDate: false,
//         showTextBlockAboveTitle: false,
//         textAboveTitle: "",
//         textAlignment: "center",
//         imageSourceType: 2,
//         title: "sample1",
//       },
//       canvasLayout: {
//         horizontalSections: [
//           {
//             layout: "oneColumn",
//             id: "1",
//             emphasis: "none",
//             columns: [
//               {
//                 id: "1",
//                 width: 12,
//                 webparts: [
//                   {
//                     id: "6f9230af-2a98-4952-b205-9ede4f9ef548",
//                     innerHtml: `<p>Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nullam et magna quis mauris tincidunt scelerisque. Nulla quis arcu euismod, ullamcorper dui et, aliquet nulla. Cras nec nisl sed mauris mattis sodales. Proin quis odio sit amet nisl dapibus iaculis sit amet id sem. Curabitur auctor nulla a lorem volutpat, vitae feugiat tortor congue. Nullam at viverra nulla, et maximus leo. Quisque dictum lorem nec purus tincidunt, at varius urna accumsan. Vestibulum vehicula, enim sit amet ullamcorper dictum, velit purus vestibulum nulla, non lacinia felis nulla nec tellus.</p>

//                     <p>Vestibulum finibus, velit in egestas placerat, ex mauris ultricies nulla, et suscipit mi eros a libero. Ut bibendum elit nec mi vestibulum, non pharetra justo hendrerit. Donec in augue id ligula tincidunt varius. Nulla facilisi. Integer tristique ultricies odio, vel finibus ipsum. Nam elementum lectus at orci fermentum, vitae blandit turpis aliquet. Nunc eget diam sed odio luctus congue nec a sem. Sed fermentum ipsum sit amet tincidunt scelerisque. Morbi vitae quam nec tortor suscipit placerat. Nulla in diam nec mi varius placerat.</p>

//                     <h3 style="color: #555; margin-bottom: 10px;">Key Points:</h3>
//                     <ul style="margin-bottom: 15px;">
//                         <li>Point 1: Vivamus accumsan odio vitae lectus consectetur, id accumsan turpis suscipit.</li>
//                         <li>Point 2: Duis vehicula consequat augue, sit amet consequat lectus tempor ac.</li>
//                         <li>Point 3: Aliquam nec tincidunt purus. Suspendisse eu risus vitae dui ultricies consequat.</li>
//                     </ul>

//                     <p>Phasellus at velit at metus suscipit tempus. Nulla vestibulum lacus sed libero dapibus, eget gravida est sagittis. Ut vitae massa vel odio lacinia luctus. Fusce tincidunt dapibus odio, a pharetra mi hendrerit id. Nunc eget semper leo. Vestibulum gravida vestibulum ipsum, in suscipit enim vestibulum nec. Nullam et diam et lectus lacinia maximus in ac nunc. Cras vitae libero non lorem sodales pretium. Ut ac diam nunc.</p>

//                     <p>Quisque imperdiet lacinia ex, in consequat risus dictum et. Nullam scelerisque, nunc id hendrerit dapibus, sem libero pellentesque libero, non suscipit quam mi id magna. Ut et varius ligula. Cras ullamcorper libero nisl, sit amet pretium quam malesuada non. Maecenas molestie vel ligula ac lacinia. Mauris sed elit eu erat congue volutpat. Nam pretium odio eget eleifend pharetra. Vestibulum vehicula, quam eu hendrerit elementum, enim justo lobortis ex, vel molestie sapien lorem id libero.</p>

//                     <h3 style="color: #555; margin-bottom: 10px;">Conclusion</h3>
//                     <p>Integer lacinia pulvinar tincidunt. Proin a ligula a quam fringilla tempus. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Fusce at mollis nisl. Cras luctus aliquet vehicula. Vivamus ut nisl odio. Mauris interdum purus ut sapien ultrices, eget scelerisque ligula faucibus. Nam a lectus nec eros mollis fermentum id a dui. Vestibulum sed neque sit amet nunc laoreet tristique in a justo. Donec dignissim, lorem quis fermentum consequat, elit eros viverra lorem, et volutpat tortor nisi a nunc.</p>`,
//                   },
//                 ],
//               },
//             ],
//           },
//         ],
//       },
//     };

//     let siteID = "787dcd94-d34c-46db-beb6-a4cae089ad8a";
//     this.props.context.msGraphClientFactory
//       .getClient("3")
//       .then((client: MSGraphClientV3): void => {
//         client
//           .api(`https://graph.microsoft.com/v1.0/sites/${siteID}/pages`)
//           .version("v1.0")

//           .post(baseSitePage)
//           .then((res: any) => {
//             const pageID = res.id;

//             console.log(res);
//             this.props.context.msGraphClientFactory
//               .getClient("3")
//               .then((client: MSGraphClientV3): void => {
//                 client
//                   .api(
//                     `https://graph.microsoft.com/v1.0/sites/${siteID}/pages/${pageID}/microsoft.graph.sitePage/publish`
//                   )
//                   .version("v1.0")
//                   .post({})
//                   .then(() => {
//                     alert("Page published");
//                   });
//               });
//           })
//           .catch((err: any) => {
//             alert("Error" + err);
//             console.log("Error", err);
//           });
//       });
//   };
