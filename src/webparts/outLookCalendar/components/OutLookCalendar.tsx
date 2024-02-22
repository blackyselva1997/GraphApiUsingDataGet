import * as React from "react";
import * as moment from "moment";
import "./Styles.css";
import { useState, useEffect } from "react";
import { MSGraphClient } from "@microsoft/sp-http";
import { Calendar } from "@fullcalendar/core";
import interactionPlugin from "@fullcalendar/interaction";
import dayGridPlugin from "@fullcalendar/daygrid";
import timeGridPlugin from "@fullcalendar/timegrid";
import listPlugin from "@fullcalendar/list";
import bootstrap5Plugin from "@fullcalendar/bootstrap5";
import styles from "./OutLookCalendar.module.scss";
import {
  DetailsList,
  PrimaryButton,
  SelectionMode,
  TooltipHost,
} from "office-ui-fabric-react";

interface IEvent {
  title: string;
  start: string;
  end: string;
}

interface IData {
  ID: string;
  DisplayName: string;
  Mail: string;
  UserPrincipalName: string;
}

// Variables
let ObjectID: string = "";
let GroupDisplayName: string = "TestGroupCalander";

const OutLookCalendar = (props: any) => {
  // Styles
  const detailListStyle = {
    root: {
      ".ms-DetailsHeader-cell": {
        background: "white !important",
        color: "#323130 !important",
        fontWeight: "bold",
      },
      ".ms-DetailsRow-cell": {
        display: "flex",
        alignItems: "center",
        justifyContent: "start",
      },
    },
  };

  // Column
  const renderColumn = (
    fieldName: string,
    label: string,
    minWidth: any,
    maxWidth: any
  ) => {
    return {
      key: fieldName,
      name: label,
      fieldName: fieldName,
      minWidth: minWidth,
      maxWidth: maxWidth,
      onRender: (item: IData) => {
        let fieldValue = item[fieldName];

        return (
          <p
            title={fieldValue}
            className={"DetailslistText"}
            style={{
              width: "100%",
              textAlign: "left",
            }}
          >
            {fieldValue ? fieldValue : "-"}
          </p>
        );
      },
    };
  };

  // Define your columns using the reusable function
  // function call 1. fieldName.2. label.3. minWidth.4. maxWidth.
  const MainColumn = [
    renderColumn("id", "ID", 200, 245),
    renderColumn("displayName", "Display Name", 100, 190),
    renderColumn("mail", "Email", 200, 265),
    renderColumn("userPrincipalName", "Principal Name", 300, 430),
  ];

  // Use States
  const [AllUsers, setAllUsers] = useState([]);
  const [calendarVisible, setCalendarVisible] = useState({
    MyEvent: true,
    GroupEvent: false,
    AllUsers: false,
  });

  // Error Functions
  const GetErrorFunctions = (type: string, error: any) => {
    console.log(type, error);
  };

  // My Events get
  const GetAllFunctions = () => {
    props.context._msGraphClientFactory
      .getClient()
      .then((graphClient: MSGraphClient) => {
        graphClient
          .api("me/events")
          .filter(
            "start/datetime ge '" +
              "2020-01-01T00:00:00.0000000" +
              "' and end/datetime le '" +
              "2024-09-20T10:00:00.0000000" +
              "'"
          )
          .top(999)
          .get()
          .then((meEvents: any) => {
            let tempTask = [];

            if (meEvents.value.length) {
              meEvents.value.forEach((task: any) => {
                var sdate =
                  moment(task.start.dateTime).format("YYYY-MM-DD") +
                  "T" +
                  moment(task.start.dateTime).format("HH:mm") +
                  ":00";
                var edate =
                  moment(task.end.dateTime).format("YYYY-MM-DD") +
                  "T" +
                  moment(task.end.dateTime).format("HH:mm") +
                  ":00";
                tempTask.push({
                  id: task.id,
                  Title: task.subject,
                  StartDate: sdate,
                  EndDate: edate,
                  display: "block",
                  description: task.bodyPreview,
                  ColorId: 89,
                  allDayEventCheck: task.isAllDay,
                  eventType: "outlook",
                });
              });
              getEvents([...tempTask]);
            }
          })
          .catch((err: any) => {
            GetErrorFunctions("My Events error ==> ", err);
          });
      });
  };

  // My and Groups Events get
  const getGroupsEvents = () => {
    props.context._msGraphClientFactory
      .getClient()
      .then((graphClient: MSGraphClient) => {
        graphClient
          .api("groups")
          .filter("displayName eq '" + `${GroupDisplayName}` + "'")
          .top(999)
          .get()
          .then(async (res: any) => {
            let tempGroupTask = [];
            ObjectID = res.value[0].id;

            if (res.value.length) {
              await graphClient
                .api(`groups/${ObjectID}/calendar/events`)
                .get()
                .then((res: any) => {
                  var sdate =
                    moment(res.value[0].start.dateTime).format("YYYY-MM-DD") +
                    "T" +
                    moment(res.value[0].start.dateTime).format("HH:mm") +
                    ":00";
                  var edate =
                    moment(res.value[0].end.dateTime).format("YYYY-MM-DD") +
                    "T" +
                    moment(res.value[0].end.dateTime).format("HH:mm") +
                    ":00";
                  tempGroupTask.push({
                    id: res.value[0].id,
                    Title: res.value[0].subject,
                    StartDate: sdate,
                    EndDate: edate,
                  });
                });
            }
            setCalendarVisible({
              AllUsers: false,
              GroupEvent: true,
              MyEvent: false,
            });
            getEvents([...tempGroupTask]);
          })
          .catch((err: any) => {
            GetErrorFunctions("Groups Events error ==> ", err);
          });
      });
  };

  // All Users get
  const GetAllUsers = () => {
    props.context._msGraphClientFactory
      .getClient()
      .then((graphClients: MSGraphClient) => {
        graphClients
          .api("users")
          .top(100)
          .get()
          .then((users: any) => {
            let outLookDatas = [];
            if (users) {
              for (let i = 0; users.value.length > i; i++) {
                outLookDatas.push(users.value[i]);
              }
              let nextLink = users["@odata.nextLink"];

              if (nextLink)
                GetPendingUsers(nextLink, graphClients, outLookDatas);
              else {
                setAllUsers([...outLookDatas]);
              }
            }
          })
          .catch((err: any) => {
            GetErrorFunctions("Users error ==> ", err);
          });
      });
  };

  // Pending All Users get
  const GetPendingUsers = (
    nextLinkSkip: any,
    graphClients: MSGraphClient,
    outLookDatas: any[]
  ) => {
    graphClients.api(nextLinkSkip).get((err, response) => {
      if (response.value) {
        for (let i = 0; response.value.length > i; i++) {
          outLookDatas.push(response.value[i]);
        }
      }
      let nextLink = response["@odata.nextLink"];

      if (nextLink) {
        GetPendingUsers(nextLink, graphClients, outLookDatas);
      } else {
        setAllUsers([...outLookDatas]);
      }
    });
  };

  console.log("All USers", AllUsers);

  // Calendar data bind
  const getEvents = (items: any) => {
    let _calendarData: IEvent[] = [];

    items.forEach((item: any) => {
      const formattedStartDate = new Date(item.StartDate).toISOString();
      const formattedEndDate = new Date(item.EndDate).toISOString();
      if (calendarVisible) {
        _calendarData.push({
          title: item.Title,
          start: formattedStartDate,
          end: formattedEndDate,
        });
      } else {
        _calendarData.push({
          title: item.Title,
          start: formattedStartDate,
          end: formattedEndDate,
        });
      }
    });

    BindCalender(_calendarData);
  };

  const BindCalender = (data: any) => {
    let calendarEl = document.getElementById("myCalendar");
    let _Calendar = new Calendar(calendarEl, {
      plugins: [
        interactionPlugin,
        dayGridPlugin,
        timeGridPlugin,
        listPlugin,
        bootstrap5Plugin,
      ],
      selectable: true,
      buttonText: {
        today: "Today",
        dayGridMonth: "Month",
        dayGridWeek: "Week",
        timeGridDay: "Day",
      },
      headerToolbar: {
        left: "today prevYear prev next nextYear",
        center: "title",
        right: "dayGridMonth dayGridWeek timeGridDay",
      },
      initialDate: new Date(),
      events: data,
      height: "auto",
      displayEventTime: false,
      weekends: true,
      dayMaxEventRows: true,
      views: {
        dayGrid: {
          dayMaxEventRows: 4,
        },
        timeGridDay: {
          slotEventOverlap: false,
        },
      },
      dateClick: function (arg: any) {},
      eventMouseEnter: function (info: any) {
        info.el.setAttribute("title", info.event.title);
      },
    });
    _Calendar.render();
    _Calendar.updateSize();
  };

  useEffect(() => {
    GetAllFunctions();
  }, []);

  return (
    <>
      <div style={{ width: "75%", height: "79vh", margin: "auto" }}>
        <div className={styles.HeadButtons}>
          <PrimaryButton
            onClick={() => {
              setCalendarVisible({
                AllUsers: false,
                GroupEvent: false,
                MyEvent: true,
              });
              GetAllFunctions();
            }}
          >
            My Event Calendar
          </PrimaryButton>
          <PrimaryButton
            onClick={() => {
              getGroupsEvents();
            }}
          >
            Group Event Calendar
          </PrimaryButton>
          <PrimaryButton
            onClick={() => {
              setCalendarVisible({
                AllUsers: true,
                GroupEvent: false,
                MyEvent: false,
              });
              GetAllUsers();
            }}
          >
            All Users
          </PrimaryButton>
        </div>
        {calendarVisible.MyEvent && (
          <div>
            <h1>My Event's</h1>
            <div
              id="myCalendar"
              style={{
                visibility: "visible",
                height: 0,
                textTransform: "capitalize",
              }}
            ></div>
          </div>
        )}
        {calendarVisible.GroupEvent && (
          <div>
            <h1>Group Event's</h1>
            <div
              id="myCalendar"
              style={{
                visibility: "visible",
                height: 0,
                textTransform: "capitalize",
              }}
            ></div>
          </div>
        )}
        {calendarVisible.AllUsers && (
          <div>
            <h1>All Users</h1>
            <div style={{ width: "100%" }}>
              <DetailsList
                items={[...AllUsers]}
                columns={MainColumn.map((column) => ({
                  ...column,
                  onRenderHeader: () => {
                    return (
                      <TooltipHost
                        content={column.name}
                        calloutProps={{ gapSpace: 0 }}
                      >
                        {column.name}
                      </TooltipHost>
                    );
                  },
                }))}
                selectionMode={SelectionMode.none}
                styles={detailListStyle}
              />
            </div>
          </div>
        )}
      </div>
    </>
  );
};

export default OutLookCalendar;
