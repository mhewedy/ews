package ews

import (
	"bytes"
	"crypto/tls"
	"fmt"
	"io/ioutil"
	"net/http"
	"net/http/httputil"
	"os"

	"github.com/Azure/go-ntlmssp"
)

const (
	soapStart = `<?xml version="1.0" encoding="utf-8" ?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
		xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" 
		xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" 
		xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  		<soap:Header>
    		<t:RequestServerVersion Version="Exchange2013_SP1" />
  		</soap:Header>
  		<soap:Body>
`
	soapEnd = `
</soap:Body></soap:Envelope>`
)

type Config struct {
	Dump    bool
	NTLM    bool
	SkipTLS bool
}

type Client interface {
	SendAndReceive(body []byte) ([]byte, error)
	GetEWSAddr() string
	GetUsername() string
}

type client struct {
	EWSAddr  string
	Username string
	Password string
	config   *Config
}

func (c *client) GetEWSAddr() string {
	return c.EWSAddr
}

func (c *client) GetUsername() string {
	return c.Username
}

func NewClient(ewsAddr, username, password string, config *Config) Client {
	return &client{
		EWSAddr:  ewsAddr,
		Username: username,
		Password: password,
		config:   config,
	}
}

func (c *client) SendAndReceive(body []byte) ([]byte, error) {

	bb := []byte(soapStart)
	bb = append(bb, body...)
	bb = append(bb, soapEnd...)

	os.WriteFile("./send.xml", bb, 0755)

	req, err := http.NewRequest("POST", c.EWSAddr, bytes.NewReader(bb))
	if err != nil {
		return nil, err
	}
	defer req.Body.Close()
	logRequest(c, req)

	req.SetBasicAuth(c.Username, c.Password)
	req.Header.Set("Content-Type", "text/xml")

	client := &http.Client{
		CheckRedirect: func(req *http.Request, via []*http.Request) error {
			return http.ErrUseLastResponse
		},
	}
	applyConfig(c.config, client)

	resp, err := client.Do(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()
	logResponse(c, resp)

	if resp.StatusCode != http.StatusOK {
		return nil, NewError(resp)
	}

	respBytes, err := ioutil.ReadAll(resp.Body)
	if err != nil {
		return nil, err
	}

	os.WriteFile("./recv.xml", respBytes, 0755)
	return respBytes, err
}

func applyConfig(config *Config, client *http.Client) {
	if config.NTLM {
		client.Transport = ntlmssp.Negotiator{
			RoundTripper: &http.Transport{
				TLSNextProto:    map[string]func(authority string, c *tls.Conn) http.RoundTripper{},
				TLSClientConfig: &tls.Config{InsecureSkipVerify: config.SkipTLS},
			},
		}
	}
	if config.SkipTLS {
		http.DefaultTransport.(*http.Transport).TLSClientConfig = &tls.Config{InsecureSkipVerify: true}
	}
}

func logRequest(c *client, req *http.Request) {
	if c.config != nil && c.config.Dump {
		dump, err := httputil.DumpRequestOut(req, true)
		if err != nil {
			fmt.Println(err)
		}
		fmt.Printf("Request:\n%v\n----\n", string(dump))
	}
}

func logResponse(c *client, resp *http.Response) {
	if c.config != nil && c.config.Dump {
		dump, err := httputil.DumpResponse(resp, true)
		if err != nil {
			fmt.Println(err)
		}
		fmt.Printf("Response:\n%v\n----\n", string(dump))
	}
}
