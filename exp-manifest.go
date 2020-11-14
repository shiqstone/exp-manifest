package main

import (
	"flag"
	"fmt"
	"golang.org/x/text/encoding/simplifiedchinese"
	"io/ioutil"
	"log"
	"os"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/BurntSushi/toml"
)

type Config struct {
	Unit []string `toml:"unit"`
	Path string  `toml:"path"`
	OutputPath string `toml:"output_path"`
}

type Item struct {
	Name string
	Unit string
	Quantity int
	Price float64
	Total float64
	Remark string
}

func main() {
	path := flag.String("p", "", "path of uno list path")
	configFile := flag.String("c", "./conf/conf.toml", "config file")
	title := flag.String("t", "", "title of export excel name")
	flag.Parse()

	conf := new(Config)
	_, err := toml.DecodeFile(*configFile, conf)
	if err != nil {
		log.Println("failed to decode config file", configFile, err)
		return
	}

	if *path == "" {
		*path = conf.Path
	}

	data := make([]*Item, 0)
	odstr := ReadFile(*path)
	odstr = strings.Replace(odstr, "\r\n", "", -1)
	odstr = strings.Replace(odstr, "，", "", -1)
	odstr = strings.Replace(odstr, "。", "", -1)
	odstr = strings.Replace(odstr, ",", "", -1)
	if !isUtf8([]byte(odstr)) {
		utfdata, _ := simplifiedchinese.GBK.NewDecoder().Bytes([]byte(odstr))
		odstr = string(utfdata)
	}

	lines := strings.Split(odstr, "元")
	log.Println(lines)
	re := regexp.MustCompile("([\u4e00-\u9fa5]+)([\\d]+)([\u4e00-\u9fa5]+)([\\d]+\\.?[\\d]?)")
	// re := regexp.MustCompile("(.+)([\\d]+)(斤|包|袋|个|盒|瓶|桶|箱|板|排)([\\d]+\\.?[\\d]?)")
	for _, line := range lines {
		if len(line) == 0 {
			continue
		}
		
		log.Println(line)
		params := re.FindStringSubmatch(line)
		log.Println(params)
		item := new(Item)
		item.Name = params[1]
		item.Unit = params[3]
		item.Quantity, _ = strconv.Atoi(params[2])
		item.Price, err = strconv.ParseFloat(params[4], 32/64)
		item.Total = item.Price * float64(item.Quantity)
		data = append(data, item)
	}

	if len(data) == 0 {
		return
	}

	f := excelize.NewFile()
	// 创建一个工作表
	index := f.NewSheet("Sheet1")
	// 名称	单位	重量/数量	单价	总价	备注
	f.SetCellValue("Sheet1", "A1", "名称")
	f.SetCellValue("Sheet1", "B1", "单位")
	f.SetCellValue("Sheet1", "C1", "重量/数量")
	f.SetCellValue("Sheet1", "D1", "单价")
	f.SetCellValue("Sheet1", "E1", "总价")
	f.SetCellValue("Sheet1", "F1", "备注")

	line := 2
	total := 0.0
	for _, item := range data {
		f.SetCellValue("Sheet1", "A"+strconv.Itoa(line), item.Name)
		f.SetCellValue("Sheet1", "B"+strconv.Itoa(line), item.Unit)
		f.SetCellValue("Sheet1", "C"+strconv.Itoa(line), item.Quantity)
		f.SetCellValue("Sheet1", "D"+strconv.Itoa(line), item.Price)
		f.SetCellValue("Sheet1", "E"+strconv.Itoa(line), item.Total)
		f.SetCellValue("Sheet1", "F"+strconv.Itoa(line), item.Remark)
	
		total += item.Total
		line++
	}
	// 设置工作簿的默认工作表
	f.SetActiveSheet(index)

	f.SetCellValue("Sheet1", "A"+strconv.Itoa(line), "合计")
	f.SetCellValue("Sheet1", "B"+strconv.Itoa(line), "")
	f.SetCellValue("Sheet1", "C"+strconv.Itoa(line), "")
	f.SetCellValue("Sheet1", "D"+strconv.Itoa(line), "")
	f.SetCellValue("Sheet1", "E"+strconv.Itoa(line), total)
	f.SetCellValue("Sheet1", "F"+strconv.Itoa(line), "")
	
	f.MergeCell("Sheet1", "A"+strconv.Itoa(line), "D"+strconv.Itoa(line))


	var fname string
	if *title == "" {
		fname = time.Now().Format("20060102150405") + ".xlsx"
	} else {
		fname = *title
	}
	expName := conf.OutputPath + fname
	err = f.SaveAs(expName)
    if err != nil {
        log.Println(err)
    }
}

func ReadFile(path string)  (string){
    f, err := os.Open(path)
    if err != nil {
        log.Println("read file fail", err)
        return ""
    }
    defer f.Close()

    fd, err := ioutil.ReadAll(f)
    if err != nil {
        log.Println("read to fd fail", err)
        return ""
    }

    return string(fd)
}

func ParseBillData(origin string) []Item {
	return nil
}

func preNUm(data byte) int {
    str := fmt.Sprintf("%b", data)
    var i int = 0
    for i < len(str) {
        if str[i] != '1' {
            break
        }
        i++
    }
    return i
}
func isUtf8(data []byte) bool {
    for i := 0; i < len(data);  {
        if data[i] & 0x80 == 0x00 {
            // 0XXX_XXXX
            i++
            continue
        } else if num := preNUm(data[i]); num > 2 {
            // 110X_XXXX 10XX_XXXX
            // 1110_XXXX 10XX_XXXX 10XX_XXXX
            // 1111_0XXX 10XX_XXXX 10XX_XXXX 10XX_XXXX
            // 1111_10XX 10XX_XXXX 10XX_XXXX 10XX_XXXX 10XX_XXXX
            // 1111_110X 10XX_XXXX 10XX_XXXX 10XX_XXXX 10XX_XXXX 10XX_XXXX
            // preNUm() 返回首个字节的8个bits中首个0bit前面1bit的个数，该数量也是该字符所使用的字节数
            i++
            for j := 0; j < num - 1; j++ {
                //判断后面的 num - 1 个字节是不是都是10开头
                if data[i] & 0xc0 != 0x80 {
                    return false
                }
                i++
            }
        } else  {
            //其他情况说明不是utf-8
            return false
        }
    }
    return true
}