package com.bishe.studentsystem.service;

import com.bishe.studentsystem.dao.*;
import com.bishe.studentsystem.dao.UserDao;
import com.bishe.studentsystem.model.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.HashMap;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;


public class AdminDao extends HttpServlet {
	private static final long serialVersionUID = 1L;
    
	private String action;//存储操作描述
	//接收请求
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		request.setCharacterEncoding("utf-8");
		action = request.getParameter("action");
		//判断所执行操作
		switch (action) {
		//用户操作
		case "query_all_user":
			query_all_user(request, response);break;
		case "insert_user":
			insert_user(request,response);break;
		case "delete_user":
			delete_user(request, response);break;
		case "alter_user":
			alter_user(request, response);break;
		//学生操作
		case "query_all_student":
			query_all_student(request, response);break;
		case "insert_student":
			insert_student(request, response);break;	
		case "delete_student":
			delete_student(request, response);break;
		case "alter_student":
			alter_student(request, response);break;
		case "download_file":
			download_file(request,response);break;
		//课程操作
		case "course_avg":
			course_avg(request, response);break;
		case "fail_rate":
			fail_rate(request, response);break;
		case "course_ranking":
			course_ranking(request, response);break;
		case "query_all_course":
			query_all_course(request, response);break;
		case "insert_course":
			insert_course(request, response);break;
		case "delete_course":
			delete_course(request, response);break;
		case "alter_course":
			alter_course(request, response);break;
		//成绩操作
		case "query_all_sc":
			query_all_sc(request, response);break;
		case "query_one_sc":
			query_one_sc(request,response);break;
		case "insert_sc":
			insert_sc(request, response);break;
		case "delete_sc":
			delete_sc(request, response);break;
		case "alter_sc":
			alter_sc(request, response);break;
		default:
			break;
		}
	}
	/*-------------------------------- 用户 -----------------------------------*/
	//查询所有用户
	protected void query_all_user(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		response.setContentType("text/html;charset=utf-8");
		UserDao userDao = new UserDao();
		
		ArrayList<User> results = userDao.query_all_user();
		PrintWriter out = response.getWriter();
		//输出结果
		if(results != null){
			out.write("<div class='all'>");
			out.write("<div><span>用户名</span><span>密码</span><span>权限级别</span></div>");
			for(User i: results){
				out.write("<div>");
				out.write("<span>"+i.getUsername()+"</span>");
				out.write("<span>"+i.getPassword()+"</span>");
				out.write("<span>"+i.getLevel()+"</span>");
				out.write("</div>");
			}
			out.write("</div>");
		}
		
		out.flush();
		out.close();
	}
	//插入用户
	protected void insert_user(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		response.setContentType("text/html;charset=utf-8");
		String username = request.getParameter("username");
		String password = request.getParameter("password");
		String level = request.getParameter("level");
		
		int flag =  new UserDao().insert_user(username, password, level);
		String info = null;
		PrintWriter out =  response.getWriter();
		if(flag == 1){
			info = "用户插入成功！";
		}else{
			info = "错误：用户插入失败！";
		}
		out.write("<div class='error'>");
		out.write("<div>"+info+"</div>");
		out.write("</div>");
		out.flush();
		out.close();
	}
	//删除用户
	protected void delete_user(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		response.setContentType("text/html;charset=utf-8");
		String username = request.getParameter("username");
		
		int flag =  new UserDao().delete_user(username);
		String info = null;
		PrintWriter out =  response.getWriter();
		if(flag == 1){
			info = "成功删除名为"+username+"用户！";
		}else{
			info = "错误：删除用户失败！";
		}
		out.write("<div class='error'>");
		out.write("<div>"+info+"</div>");
		out.write("</div>");
		out.flush();
		out.close();
	}
	//修改用户
	protected void alter_user(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		response.setContentType("text/html;charset=utf-8");
		String username = request.getParameter("username");
		String after_username = request.getParameter("after_username");
		String after_password = request.getParameter("after_password");
		String after_level = request.getParameter("after_level");
		
		int flag =  new UserDao().alter_user(username,after_username,after_password,after_level);
		String info = null;
		PrintWriter out =  response.getWriter();
		if(flag == 1){
			info = "名为"+username+"用户信息修改成功！";
		}else{
			info = "错误：修改用户失败!";
		}
		out.write("<div class='error'>");
		out.write("<div>"+info+"</div>");
		out.write("</div>");
		out.flush();
		out.close();
	}

	/*-------------------------------- 学生-----------------------------------*/
	// 查询所有学生
	protected void query_all_student(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		response.setContentType("text/html;charset=utf-8");
		ArrayList<Student> results = new StudentDao().query_all_student();
		PrintWriter out = response.getWriter();
		// 输出结果
		if (results != null) {
			out.write("<div class='all'>");
			out.write("<div><span>学号</span><span>姓名</span><span>性别</span><span>年龄</span><span>所在班级编号</span></div>");
			for (Student i : results) {
				out.write("<div>");
				out.write("<span>" + i.getSno() + "</span>");
				out.write("<span>" + i.getSname() + "</span>");
				out.write("<span>" + i.getSsex() + "</span>");
				out.write("<span>" + i.getSage() + "</span>");
				out.write("<span>" + i.getClno() + "</span>");
				out.write("</div>");
			}
			out.write("</div>");
		}
		out.flush();
		out.close();
	}

	// 插入学生
	protected void insert_student(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		response.setContentType("text/html;charset=utf-8");
		String sno = request.getParameter("sno");
		String sname = request.getParameter("sname");
		String ssex = request.getParameter("ssex");
		int sage = Integer.parseInt(request.getParameter("sage"));
		String clno = request.getParameter("clno");
		int flag = new StudentDao().insert_student(sno, sname, ssex, sage, clno);
		String info = null;
		PrintWriter out = response.getWriter();
		if (flag == 1) {
			info = "学生"+sname+"插入成功！";
		} else {
			info = "错误：学生插入失败！";
		}
		out.write("<div class='error'>");
		out.write("<div>" + info + "</div>");
		out.write("</div>");
		out.flush();
		out.close();
	}
	// 删除学生
	protected void delete_student(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		response.setContentType("text/html;charset=utf-8");
		String sno = request.getParameter("sno");
		int flag = new StudentDao().delete_student(sno);
		String info = null;
		PrintWriter out = response.getWriter();
		if (flag == 1) {
			info = "成功删除" + sno + "号学生！";
		} else {
			info = "错误：删除学生失败！";
		}
		out.write("<div class='error'>");
		out.write("<div>" + info + "</div>");
		out.write("</div>");
		out.flush();
		out.close();
	}
	// 修改学生信息
	protected void alter_student(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		response.setContentType("text/html;charset=utf-8");
		String sno = request.getParameter("sno");
		String after_sno = request.getParameter("after_sno");
		String after_sname = request.getParameter("after_sname");
		String after_ssex = request.getParameter("after_ssex");
		int after_sage = Integer.parseInt(request.getParameter("after_sage"));
		String after_clno = request.getParameter("after_clno");
		int flag = new StudentDao().alter_class(sno, after_sno, after_sname, after_ssex, after_sage, after_clno);
		String info = null;
		PrintWriter out = response.getWriter();
		if (flag == 1) {
			info = "学生"+sno+"信息修改成功！";
		} else {
			info = "错误：修改学生信息失败!";
		}
		out.write("<div class='error'>");
		out.write("<div>" + info + "</div>");
		out.write("</div>");
		out.flush();
		out.close();
	}

	/*-------------------------------- 课程 -----------------------------------*/
	//查询课程平均分
	protected void course_avg(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {

		response.setContentType("text/html;charset=utf-8");
		ArrayList<Course_avg> results = new CourseDao().course_avg();
		PrintWriter out = response.getWriter();
		if(results != null){
			//输出结果
			if(results != null){
				out.write("<div class='all'>");
				out.write("<div><span>课程号</span><span>课程名称</span><span>平均分</span></div>");
				for(Course_avg i:results) {
					out.write("<div>");
					out.write("<span>"+i.getCno()+"</span>");
					out.write("<span>"+i.getCname()+"</span>");
					out.write("<span>"+i.getAvg()+"</span>");
					out.write("</div>");
				}
				out.write("</div>");
			}
		}
		out.flush();
		out.close();
	}
	//查询课程不及格率
	protected void fail_rate(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		response.setContentType("text/html;charset=utf-8");
		ArrayList<Course_fail_rate> results = new CourseDao().fail_rate();
		PrintWriter out = response.getWriter();
		// 输出结果
		if (results != null) {
			out.write("<div class='all'>");
			out.write("<div><span>课程号</span><span>课程名称</span><span>不及格率</span></div>");
			for (Course_fail_rate i : results) {
				out.write("<div>");
				out.write("<span>" + i.getCno() + "</span>");
				out.write("<span>" + i.getCname() + "</span>");
				out.write("<span>" + i.getFail_rate() + "%</span>");
				out.write("</div>");
			}
			out.write("</div>");
		}
		out.flush();
		out.close();
	}
	// 查询课程排名
	protected void course_ranking(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		response.setContentType("text/html;charset=utf-8");
		String Cno = request.getParameter("cno");
		ArrayList<Course_ranking> results = new CourseDao().course_ranking(Cno);
		PrintWriter out = response.getWriter();
		// 输出结果
		if (results != null) {
			out.write("<div class='all'>");
			out.write("<div><span>学号</span><span>所在系</span><span>班级名称</span><span>姓名</span><span>性别</span><span>年龄</span><span>成绩</span><span>排名</span></div>");
			int no = 1;
			for (Course_ranking i : results) {
				out.write("<div>");
				out.write("<span>" + i.getSno() + "</span>");
				out.write("<span>" + i.getDname() + "</span>");
				out.write("<span>" + i.getClname() + "</span>");
				out.write("<span>" + i.getSname() + "</span>");
				out.write("<span>" + i.getSsex() + "</span>");
				out.write("<span>" + i.getSage() + "</span>");
				out.write("<span>" + i.getGrade() + "</span>");
				out.write("<span>"+(no++)+"</span>");
				out.write("</div>");
			}
			out.write("</div>");
		}
		out.flush();
		out.close();
	}
	//查询所有课程
	protected void query_all_course(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		response.setContentType("text/html;charset=utf-8");
		ArrayList<Course> results = new CourseDao().query_all_course();
		PrintWriter out = response.getWriter();
		if(results != null){
			//输出结果
			if(results != null){
				out.write("<div class='all'>");
				out.write("<div><span>课程号</span><span>课程名称</span><span>执教老师</span><span>学分</span></div>");
				for(Course i:results) {
					out.write("<div>");
					out.write("<span>"+i.getCno()+"</span>");
					out.write("<span>"+i.getCname()+"</span>");
					out.write("<span>"+i.getCteacher()+"</span>");
					out.write("<span>"+i.getCcredit()+"</span>");
					out.write("</div>");
				}
				out.write("</div>");
			}
		}
		out.flush();
		out.close();
	}
	//插入课程
	protected void insert_course(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		response.setContentType("text/html;charset=utf-8");
		String Cno = request.getParameter("cno");
		String Cname = request.getParameter("cname");
		String Cteacher = request.getParameter("cteacher");
		int Ccredit = Integer.parseInt(request.getParameter("ccredit"));
		int flag =  new CourseDao().insert_course(Cno, Cname, Cteacher, Ccredit);
		String info = null;
		PrintWriter out =  response.getWriter();
		if(flag == 1){
			info = "课程"+Cname+"插入成功！";
		}else{
			info = "错误：课程插入失败！";
		}
		out.write("<div class='error'>");
		out.write("<div>"+info+"</div>");
		out.write("</div>");
		out.flush();
		out.close();
	}
	//删除课程
	protected void delete_course(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		response.setContentType("text/html;charset=utf-8");
		String cno = request.getParameter("cno");
		int flag =  new CourseDao().delete_course(cno);
		String info = null;
		PrintWriter out =  response.getWriter();
		if(flag == 1){
			info = "成功删除"+cno+"课程！";
		}else{
			info = "错误：删除课程失败！";
		}
		out.write("<div class='error'>");
		out.write("<div>"+info+"</div>");
		out.write("</div>");
		out.flush();
		out.close();
	}
	//修改课程信息
	protected void alter_course(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		response.setContentType("text/html;charset=utf-8");
		
		String cno = request.getParameter("cno");
		String after_cno = request.getParameter("after_cno");
		String after_cname = request.getParameter("after_cname");
		String after_cteacher = request.getParameter("after_cteacher");
		double after_ccredit = Double.parseDouble(request.getParameter("after_ccredit"));
		int flag = new CourseDao().alter_course(cno, after_cno, after_cname, after_cteacher, after_ccredit);
		String info = null;
		PrintWriter out = response.getWriter();
		if (flag == 1) {
			info = cno + "号课程修改成功！";
		} else {
			info = "错误：修改课程信息失败!";
		}
		out.write("<div class='error'>");
		out.write("<div>" + info + "</div>");
		out.write("</div>");
		out.flush();
		out.close();
	}
	protected void download_file(HttpServletRequest request, HttpServletResponse response) throws IOException {
		// 设置响应头，告知浏览器这是一个Excel文件，并且指定文件名以及编码格式等信息
		response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		response.setHeader("Content-Disposition", "attachment;filename=" + java.net.URLEncoder.encode("student_course_grades.xlsx", "UTF-8"));

		ArrayList<SC> queryResults = new SCDao().query_all_sc();

		// 创建一个新的Excel工作簿（对应一个.xlsx文件）
		Workbook workbook = new XSSFWorkbook();

		// 创建一个工作表
		Sheet sheet = workbook.createSheet("Student_Course_Grades");

		// 创建表头行
		String[] headers = {"学号", "姓名", "性别", "年龄", "课程编号", "课程名称", "成绩"};
		Row headerRow = sheet.createRow(0);
		for (int i = 0; i < headers.length; i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellValue(headers[i]);
		}

		// 填充数据行，遍历查询结果集填充到Excel表格中
		for (int rowIndex = 0; rowIndex < queryResults.size(); rowIndex++) {
			SC sc = queryResults.get(rowIndex);
			Row dataRow = sheet.createRow(rowIndex + 1);
			dataRow.createCell(0).setCellValue(sc.getSno());
			dataRow.createCell(1).setCellValue(sc.getSname());
			dataRow.createCell(2).setCellValue(sc.getSsex());
			dataRow.createCell(3).setCellValue(sc.getSage());
			dataRow.createCell(4).setCellValue(sc.getCno());
			dataRow.createCell(5).setCellValue(sc.getCname());
			dataRow.createCell(6).setCellValue(sc.getGrade());
		}


		// 获取输出流，用于将Excel文件内容输出到客户端
		OutputStream outputStream = response.getOutputStream();
		// 将工作簿写入输出流
		workbook.write(outputStream);
		// 关闭工作簿和输出流
		workbook.close();
		outputStream.close();
	}

	/*-------------------------------- 成绩-----------------------------------*/
	// 查询所有成绩表
	protected void query_all_sc(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		response.setContentType("text/html;charset=utf-8");
		ArrayList<SC> results = new SCDao().query_all_sc();
		PrintWriter out = response.getWriter();
		// 输出结果
		if (results != null) {
			out.write("<div id='all_sc' class='all'>");
			out.write("    <style>\n" +
					"        button {\n" +
					"            background-color: #007bff; /* 蓝色背景色，你可以替换成喜欢的颜色 */\n" +
					"            color: white; /* 文字颜色为白色 */\n" +
					"            border: 2px solid #007bff; /* 2px 宽的同色边框 */\n" +
					"            padding: 10px 20px; /* 上下 10px 左右 20px 的内边距 */\n" +
					"            border-radius: 5px; /* 5px 的圆角 */\n" +
					"            cursor: pointer; /* 鼠标悬停时显示手型指针，体现可点击性 */\n" +
					"        }\n" +
					"\n" +
					"        button:hover {\n" +
					"            background-color: white; /* 鼠标悬停时背景变为白色 */\n" +
					"            color: #007bff; /* 文字颜色变为按钮原本的背景色 */\n" +
					"            border-color: #007bff; /* 边框颜色保持不变 */\n" +
					"        }\n" +
					"    </style>");
			out.write("<button onclick=\"download_file()\">导出excel</button>");
			out.write("<div><span>学号</span><span>姓名</span><span>性别</span><span>年龄</span><span>课程号</span><span>课程名称</span><span>分数</span></div>");
			for (SC i : results) {
				out.write("<div>");
				out.write("<span>" + i.getSno() + "</span>");
				out.write("<span>" + i.getSname() + "</span>");
				out.write("<span>" + i.getSsex() + "</span>");
				out.write("<span>" + i.getSage() + "</span>");
				out.write("<span>" + i.getCno() + "</span>");
				out.write("<span>" + i.getCname() + "</span>");
				out.write("<span>" + i.getGrade() + "</span>");
				out.write("</div>");
			}
			out.write("</div>");
		}
		out.flush();
		out.close();
	}


	/*-------------------------------- 成绩-----------------------------------*/
	// 查询个人成绩
	protected void query_one_sc(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		String sno = request.getParameter("sno");
		String sname = request.getParameter("sname");

		//存在学号用学号查询
		if (!sno.equals("")) {
			ArrayList<HashMap> results = new SCDao().query_one_sc_by_sno(sno);
			response.setContentType("text/html;charset=utf-8");
			PrintWriter out = response.getWriter();
			// 输出结果
			if (results != null) {
				out.write("<div id='one_sc' class='all'>");
				out.write("<div><span>学号</span><span>姓名</span><span>课程号</span><span>课程名称</span><span>分数</span></div>");
				for (HashMap i : results) {
					out.write("<div>");
					out.write("<span>" + i.get("sno") + "</span>");
					out.write("<span>" + i.get("sname") + "</span>");
					out.write("<span>" + i.get("cno") + "</span>");
					out.write("<span>" + i.get("cname") + "</span>");
					out.write("<span>" + i.get("grade") + "</span>");

					out.write("</div>");
				}
				out.write("</div>");
			}
			out.flush();
			out.close();
		} else if (sname != null) {
			ArrayList<HashMap> results = new SCDao().query_one_sc_by_sname(sname);
			response.setContentType("text/html;charset=utf-8");
			PrintWriter out = response.getWriter();
			// 输出结果
			if (results != null) {
				out.write("<div id='one_sc' class='all'>");
				out.write("<div><span>学号</span><span>姓名</span><span>课程号</span><span>课程名称</span><span>分数</span></div>");
				for (HashMap i : results) {
					out.write("<div>");
					out.write("<span>" + i.get("sno") + "</span>");
					out.write("<span>" + i.get("sname") + "</span>");
					out.write("<span>" + i.get("cno") + "</span>");
					out.write("<span>" + i.get("cname") + "</span>");
					out.write("<span>" + i.get("grade") + "</span>");

					out.write("</div>");
				}
				out.write("</div>");
			}
			out.flush();
			out.close();

		}


		response.setContentType("text/html;charset=utf-8");
		ArrayList<SC> results = new SCDao().query_all_sc();
		PrintWriter out = response.getWriter();
		// 输出结果
		if (results != null) {
			out.write("<div id='all_sc' class='all'>");
			out.write("<div><span>学号</span><span>姓名</span><span>性别</span><span>年龄</span><span>课程号</span><span>课程名称</span><span>分数</span></div>");
			for (SC i : results) {
				out.write("<div>");
				out.write("<span>" + i.getSno() + "</span>");
				out.write("<span>" + i.getSname() + "</span>");
				out.write("<span>" + i.getSsex() + "</span>");
				out.write("<span>" + i.getSage() + "</span>");
				out.write("<span>" + i.getCno() + "</span>");
				out.write("<span>" + i.getCname() + "</span>");
				out.write("<span>" + i.getGrade() + "</span>");
				out.write("</div>");
			}
			out.write("</div>");
		}
		out.flush();
		out.close();
	}


	// 插入一条成绩记录
	protected void insert_sc(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		response.setContentType("text/html;charset=utf-8");
		String sno = request.getParameter("sno");
		String cno = request.getParameter("cno");
		double grade = Double.parseDouble(request.getParameter("grade"));
		int flag = new SCDao().insert_sc(sno, cno, grade);
		String info = null;
		PrintWriter out = response.getWriter();
		if (flag == 1) {
			info = sno + "号学生" + cno + "号课程成绩"+grade+"插入成功！";
		} else {
			info = "错误：成绩信息插入失败！";
		}
		out.write("<div class='error'>");
		out.write("<div>" + info + "</div>");
		out.write("</div>");
		out.flush();
		out.close();
	}
	// 删除成绩记录
	protected void delete_sc(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		response.setContentType("text/html;charset=utf-8");
		String sno = request.getParameter("sno");
		String cno = request.getParameter("cno");
		int flag = new SCDao().delete_sc(sno, cno);
		String info = null;
		PrintWriter out = response.getWriter();
		if (flag == 1) {
			info = "成功删除" + sno + "号学生"+cno+"号课程成绩！";
		} else {
			info = "错误：删除成绩信息失败！";
		}
		out.write("<div class='error'>");
		out.write("<div>" + info + "</div>");
		out.write("</div>");
		out.flush();
		out.close();
	}
	// 修改成绩信息
	protected void alter_sc(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		response.setContentType("text/html;charset=utf-8");
		String sno = request.getParameter("sno");
		String cno = request.getParameter("cno");
		double after_grade = Double.parseDouble(request.getParameter("after_grade"));
		int flag = new SCDao().alter_sc(sno, cno, after_grade);
		String info = null;
		PrintWriter out = response.getWriter();
		if (flag == 1) {
			info = sno + "号学生" + cno + "号课程成绩修改成功！";
		} else {
			info = "错误：修改学生成绩失败!";
		}
		out.write("<div class='error'>");
		out.write("<div>" + info + "</div>");
		out.write("</div>");
		out.flush();
		out.close();
	}

}
