package com.zekun.dev.exceltools.controller;

import com.zekun.dev.exceltools.model.ExceltFileModel;
import com.zekun.dev.exceltools.service.serviceimpl.CountWorkTimeGrantImpl;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

@RestController
public class UserController {

    @Autowired
    private CountWorkTimeGrantImpl countWorkTimeGrant;

    /**
     * 首页
     * @return
     */
    @RequestMapping("/index")
    public ModelAndView page(Model model){
        ModelAndView view = new ModelAndView();
        view.setViewName("index");
        view.addObject("exceltFileModel", new ExceltFileModel());
        return view;
    }


    /**
     * 跳转
     * @return
     */
    @RequestMapping("/redirect")
    public ModelAndView page2(){
        ModelAndView view = new ModelAndView();
        view.setViewName("redirect/redirect");
        return view;
    }


    /**
     *视图
     * @param exceltFileModel
     * @return
     */
    @PostMapping("/result")
    public ModelAndView greetingSubmit(@ModelAttribute ExceltFileModel exceltFileModel) {
        ModelAndView view = new ModelAndView();
        view.setViewName("result");
        String resultPath = countWorkTimeGrant.countResultFromExcel(exceltFileModel);
        view.addObject("resultPath", resultPath);
        return view;
    }

    @PostMapping("/shutdown")
    public void shutdown() {
        System.exit(0);
    }

}
