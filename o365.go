/*
 * Original work Copyright (c) 2019 Chris Hendricks
 * Modifications Copyright (c) 2019 John Floren
 */

package o365

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"net/http"
	"net/url"
	"sort"
	"strconv"
	"strings"
	"time"
)

const (
	PlanEnterprise        Plan = 0
	PlanGCCGovernment     Plan = 1
	PlanGCCHighGovernment Plan = 2
	PlanDODGovernment     Plan = 3

	enterpriseName        string = `enterprise`
	gccGovernmentName     string = `gccgovernment`
	gccHighGovernmentName string = `gcchighgovernment`
	dodGovernmentName     string = `dodgovernment`
)

type Plan int

// authInfo holds information returned by the microsoft oauth API
// (see https://docs.microsoft.com/en-us/office/office-365-management-api/get-started-with-office-365-management-apis#sample-response)
type authInfo struct {
	TokenType   string `json:"token_type"`
	ExpiresIn   string `json:"expires_in"`
	ExpiresOn   string `json:"expires_on"`
	NotBefore   string `json:"not_before"`
	Resource    string `json:"resource"`
	AccessToken string `json:"access_token"`
}

func (a *authInfo) header() string {
	return fmt.Sprintf("%s %s", a.TokenType, a.AccessToken)
}

func (a *authInfo) expired() bool {
	const expirationBuffer = 60 // extra seconds unexpired token is considered expired
	expiresOn, _ := strconv.ParseInt(a.ExpiresOn, 10, 64)
	return time.Now().Unix() > (expiresOn - expirationBuffer)
}

// O365beat configuration and state.
type O365 struct {
	done       chan struct{} // channel to initiate shutdown of main event loop
	config     O365Config
	authURL    string // oauth authentication url built from config
	apiRootURL string // api root url built from config
	httpClient *http.Client
	auth       *authInfo
}

type O365Config struct {
	Period        time.Duration
	ContentMaxAge time.Duration
	TenantDomain  string
	ClientSecret  string
	ClientID      string // aka application id
	DirectoryID   string // aka tenant id
	PlanName      string // the Subscription plan
	APITimeout    time.Duration
	ContentTypes  []string
}

var DefaultConfig = O365Config{
	Period:        60 * 5 * time.Second,
	APITimeout:    30 * time.Second,
	ContentMaxAge: (7 * 24 * 60) * time.Minute,
}

// New creates an instance of the O365 client.
func New(c O365Config) (*O365, error) {
	var p Plan
	if err := p.Parse(c.PlanName); err != nil {
		return nil, fmt.Errorf("invalid plan %w", err)
	}
	// using url.Parse seems like overkill
	loginURL := "https://login.microsoftonline.com/"
	au := loginURL + c.TenantDomain + "/oauth2/token?api-version=1.0"
	api := p.BaseURL() + c.DirectoryID + "/activity/feed/"
	cl := &http.Client{Timeout: c.APITimeout}
	var ai *authInfo

	o := &O365{
		done:       make(chan struct{}),
		config:     c,
		authURL:    au,
		apiRootURL: api,
		httpClient: cl,
		auth:       ai,
	}
	return o, nil
}

// apiRequest issues an http request with api authorization header
func (o *O365) apiRequest(verb, urlStr string, body, query, headers map[string]string) (*http.Response, error) {
	reqBody := url.Values{}
	for k, v := range body {
		reqBody.Set(k, v)
	}
	req, err := http.NewRequest(verb, urlStr, strings.NewReader(reqBody.Encode()))
	if err != nil {
		return nil, err
	}
	reqQuery := req.URL.Query() // keep querystring values from urlStr
	for k, v := range query {
		reqQuery.Set(k, v)
	}
	req.URL.RawQuery = reqQuery.Encode()
	for k, v := range headers {
		req.Header.Set(k, v)
	}
	// refresh authentication if expired
	if o.auth == nil || o.auth.expired() {
		err = o.authenticate()
		if err != nil {
			return nil, err
		}
	}
	req.Header.Set("Authorization", o.auth.header())

	res, err := o.httpClient.Do(req)
	if err != nil {
		return nil, err
	} else if res.StatusCode != 200 {
		body, err := ioutil.ReadAll(res.Body)
		err = fmt.Errorf("non-200 status during api request.\n\tnewly enabled or newly subscribed feeds can take 12 hours or more to provide data.\n\tconfirm audit log searching is enabled for the target tenancy (https://docs.microsoft.com/en-us/microsoft-365/compliance/turn-audit-log-search-on-or-off#turn-on-audit-log-search).\n\treq: %v\n\tres: %v\n\t%v", req, res, string(body))
		return nil, err
	}
	return res, nil
}

// authenticate retrieves oauth2 information using client id and client_secret for use with the API
// https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow
func (o *O365) authenticate() error {
	reqBody := url.Values{}
	reqBody.Set("grant_type", "client_credentials")
	reqBody.Set("resource", "https://manage.office.com")
	reqBody.Set("client_id", o.config.ClientID)
	reqBody.Set("client_secret", o.config.ClientSecret)
	req, err := http.NewRequest("POST", o.authURL, strings.NewReader(reqBody.Encode()))
	req.Header.Set("Content-Type", "application/x-www-form-urlencoded")
	res, err := o.httpClient.Do(req)
	if err != nil {
		return err
	} else if res.StatusCode != 200 {
		err = fmt.Errorf("non-200 status during auth.\n\treq: %v\n\tres: %v", req, res)
		return err
	}
	defer res.Body.Close()
	var ai authInfo
	err = json.NewDecoder(res.Body).Decode(&ai)
	if err == nil {
		o.auth = &ai
	}
	return err
}

// ListSubscriptions gets a collection of the current subscriptions and associated webhooks
// https://docs.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-reference#list-current-subscriptions
func (o *O365) ListSubscriptions() ([]map[string]string, error) {
	query := map[string]string{"PublisherIdentifier": o.config.DirectoryID}
	res, err := o.apiRequest("GET", o.apiRootURL+"subscriptions/list", nil, query, nil)
	if err != nil {
		return nil, err
	}
	defer res.Body.Close()

	var subs []map[string]string
	err = json.NewDecoder(res.Body).Decode(&subs)
	return subs, err
}

// Subscribe starts a subscription to the specified content type
// https://docs.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-reference#start-a-subscription
func (o *O365) Subscribe(contentType string) (map[string]interface{}, error) {
	query := map[string]string{
		"contentType":         contentType,
		"PublisherIdentifier": o.config.DirectoryID,
	}
	res, err := o.apiRequest("POST", o.apiRootURL+"subscriptions/start", nil, query, nil)
	if err != nil {
		return nil, err
	}
	defer res.Body.Close()

	var sub map[string]interface{}
	err = json.NewDecoder(res.Body).Decode(&sub)
	return sub, err
}

// EnableSubscriptions enables subscriptions for all configured contentTypes
func (o *O365) EnableSubscriptions() error {
	subscriptions, err := o.ListSubscriptions()
	if err != nil {
		return err
	}

	// add subscriptions as "disabled" if not in ListSubscription results (can return []!):
	for _, t := range o.config.ContentTypes {
		found := false
		for _, sub := range subscriptions {
			if sub["contentType"] == t {
				found = true
				break
			}
		}
		if !found {
			subscriptions = append(subscriptions, map[string]string{"contentType": t, "status": "disabled"})
		}
	}

	for _, sub := range subscriptions {
		if sub["status"] != "enabled" {
			_, err := o.Subscribe(sub["contentType"])
			if err != nil {
				return err
			}
		}
	}
	return nil
}

// listAvailableContent gets blob locations for a single content type over <=24 hour span
// (the basic primitive provided by the API)
// https://docs.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-reference#list-available-content
func (o *O365) ListAvailableContent(contentType string, start, end time.Time) ([]map[string]string, error) {
	now := time.Now()
	if now.Sub(start) > o.config.ContentMaxAge {
		start = now.Add(-o.config.ContentMaxAge)
	}
	if end.Sub(start).Hours() > 24 {
		err := fmt.Errorf("start (%v) and end (%v) must be <=24 hrs apart", start, end)
		return nil, err
	}
	if end.Before(start) {
		err := fmt.Errorf("start (%v) must be before end (%v)", start, end)
		return nil, err
	}

	dateFmt := "2006-01-02T15:04:05" // API needs UTC in this format (no "Z" suffix)
	query := map[string]string{
		"contentType":         contentType,
		"startTime":           start.UTC().Format(dateFmt),
		"endTime":             end.UTC().Format(dateFmt),
		"PublisherIdentifier": o.config.DirectoryID,
	}
	res, err := o.apiRequest("GET", o.apiRootURL+"subscriptions/content", nil, query, nil)
	if err != nil {
		return nil, err
	}

	var locs []map[string]string
	err = json.NewDecoder(res.Body).Decode(&locs)
	res.Body.Close()
	if err != nil {
		return nil, err
	}
	contentList := locs

	for res.Header.Get("NextPageUri") != "" {
		next := res.Header.Get("NextPageUri")
		res, err = o.apiRequest("GET", next, nil, nil, nil) // don't redeclare res!
		if err != nil {
			return nil, err
		}
		json.NewDecoder(res.Body).Decode(&locs)
		res.Body.Close()
		contentList = append(contentList, locs...)
	}
	return contentList, nil
}

// listAllAvailableContent gets blob locations for multiple content types and spans up to 7 days
// sorted by contentCreated timestamp (uses the listAvailableContent function)
func (o *O365) ListAllAvailableContent(start, end time.Time) ([]map[string]string, error) {
	if end.Before(start) {
		err := fmt.Errorf("start (%v) must be before end (%v)", start, end)
		return nil, err
	}

	interval := 24 * time.Hour
	var contentList []map[string]string

	// loop through intervals:
	for iStart, iEnd := start, start; iStart.Before(end); iStart = iEnd {
		iEnd = iStart.Add(interval)
		if end.Before(iEnd) {
			iEnd = end
		}

		// loop through all content types this interval:
		for _, t := range o.config.ContentTypes {
			list, err := o.ListAvailableContent(t, iStart, iEnd)
			if err != nil {
				return nil, err
			}
			contentList = append(contentList, list...)
		}
		// could start downloads here if concurrency is implemented
	}
	less := func(i, j int) bool {
		it, _ := time.Parse(time.RFC3339, contentList[i]["contentCreated"])
		jt, _ := time.Parse(time.RFC3339, contentList[j]["contentCreated"])
		return it.Before(jt)
	}
	sorted := sort.SliceIsSorted(contentList, less)
	if !sorted {
		sort.SliceStable(contentList, less)
	}
	return contentList, nil
}

// getContent gets actual content blobs
func (o *O365) GetContent(urlStr string) ([]byte, error) {
	query := map[string]string{
		"PublisherIdentifier": o.config.DirectoryID,
	}
	res, err := o.apiRequest("GET", urlStr, nil, query, nil)
	if err != nil {
		return nil, err
	}
	defer res.Body.Close()
	return ioutil.ReadAll(res.Body)
}

func (p *Plan) Parse(v string) (err error) {
	if v = strings.TrimSpace(v); v == `` {
		*p = PlanEnterprise
		return
	}
	tv := strings.Join(strings.Fields(strings.ToLower(v)), ``)
	switch tv {
	case enterpriseName:
		*p = PlanEnterprise //default is enterprise
	case gccGovernmentName:
		*p = PlanGCCGovernment
	case gccHighGovernmentName:
		*p = PlanGCCHighGovernment
	case dodGovernmentName:
		*p = PlanDODGovernment
	default:
		err = fmt.Errorf("unknown plan name %q", v)

	}
	return
}

func (p *Plan) String() string {
	switch *p {
	case PlanEnterprise:
		return `Enterprise`
	case PlanGCCGovernment:
		return `GCC Government`
	case PlanGCCHighGovernment:
		return `GCC High Government`
	case PlanDODGovernment:
		return `DOD Government`
	}
	return ``
}

func (p *Plan) BaseURL() string {
	switch *p {
	case PlanEnterprise:
		return `https://manage.office.com/api/v1.0/`
	case PlanGCCGovernment:
		return `https://manage-gcc.office.com/api/v1.0/`
	case PlanGCCHighGovernment:
		return `https://manage.office365.us/api/v1.0/`
	case PlanDODGovernment:
		return `https://manage.protection.apps.mil/api/v1.0/`
	}
	return `https://manage.office.com/api/v1.0/`
}
