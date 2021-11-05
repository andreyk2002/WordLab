function nowAsDuration(){
    return moment.duration({
        hours:   moment().hour(),
        minutes: moment().minute(),
        seconds: moment().second()
    });
}


$(".date").on("change", function() {
    let dateNonFormat = moment(this.value);
    this.setAttribute("my-date", dateNonFormat.format( this.getAttribute("my-date-format") )
    )
}).trigger("change")



